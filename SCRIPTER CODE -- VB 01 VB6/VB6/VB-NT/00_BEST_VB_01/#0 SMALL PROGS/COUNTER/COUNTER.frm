VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Interval        =   5
      Left            =   5040
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   5040
      Top             =   4920
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   5040
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   5040
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   9735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   9735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   9735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A1 As Currency, A2 As Currency, C1
Dim A3 As Currency, A4 As Double
Dim TX As Double, TV As Double, TF



Private Sub Form_Load()
TF = True
Label4.FontSize = 28
Call Timer4_Timer
End Sub

Private Sub Form_Resize()
    
    
    If Me.WindowState = vbMinimized Then Call TIMEROUT(False): Exit Sub
    If Me.WindowState = vbMaximized Then Call TIMEROUT(True): Exit Sub
    If Me.WindowState = vbNormal Then Call TIMEROUT(True): Exit Sub

End Sub
Sub TIMEROUT(TF)
    On Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            If Control.Interval > 0 Then Control.Enabled = TF
        Next
    Next
End Sub
Private Sub Label1_Click()
    TF = Not TF
    On Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            Err.Clear
            If Control.Interval > 0 Then
                If Err.Number = 0 Then
                    XZ = Control.Interval
                    Control.Interval = 0
                    Control.Enabled = False
                    Control.Enabled = TF
                    Control.Interval = XZ
               End If
            End If
        Next
    Next
Call Timer1_Timer
Call Timer2_Timer
Call Timer3_Timer
Call Timer4_Timer
Call Timer5_Timer

End Sub

Private Sub Timer1_Timer()

A1 = A1 + 1
C1 = Abs(A1 - A2)
'Label1 = Str(A1) + " -- " + Str(c1)
Label1 = A1
Label5 = C1
End Sub

Private Sub Timer2_Timer()
A2 = A2 + 1
C1 = Abs(A1 - A2)
'Label2 = Str(A2) + " -- " + Str(c1)
Label2 = A2
Label5 = C1
End Sub
Private Sub Timer3_Timer()
'Label1 A1
A3 = A3 + 1

TX = 0.0001 * Sin(A3) + 0.0001

TV = TV + TX
A4 = A4 + 1 + TV
Label3 = Int(Str(A4))
'Label4 = Format(TX, "0.00000000")

End Sub

Private Sub Timer4_Timer()
Label4 = Format(Now, "DDD MM MMM YYYY - hh:mm:sSa/p") + " - NOW"

End Sub

Private Sub Timer5_Timer()
'Label5 = C1

End Sub
