VERSION 5.00
Begin VB.Form UNLOAD_FORM_SAFE 
   Caption         =   "Form2"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   2955
      Top             =   1725
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   55000
      Left            =   2925
      Top             =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Clipboard Logger Still Loaded Not Propering Exit -- Force Terminate 1 Min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   -15
      TabIndex        =   0
      Top             =   -60
      Width           =   7590
   End
End
Attribute VB_Name = "UNLOAD_FORM_SAFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RETRY1

Private Sub Form_Activate()
    Unload Me

End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    Me.Show
    If IsIDE = True Then
        Timer1.Enabled = False
        Timer1.Interval = 3000
        Timer1.Enabled = True
        
    End If
    RETRY1 = RETRY1 + 1
    RETRY2 = RETRY2 + 1
    RETRY3 = RETRY3 + 1
    Label1.Caption = "RETRY1 IN THE SUB COUNTER" + Str(RETRY1) + " -- RETRY 2 PUBLIC COUNTER" + Str(RETRY2) + " -- RETRY 3 FORM LOAD " + Str(RETRY) + " -- Clipboard Logger Still Loaded Not Propering Exit -- Force Terminate 1 Min"
    
    If IsIDE = True And RETRY2 > 4 Then End
    If IsIDE = False And RETRY2 > 5 Then End
    
    
   
End Sub

Private Sub Label1_Click()

    'Stop
    'Exit Sub
    EXIT_TRUE = True
    
    'End

    Unload Form1
    Unload Me

    RETRY1 = RETRY1 + 1
    RETRY2 = RETRY2 + 1
    RETRY3 = RETRY3 + 1
    Label1.Caption = "RETRY1" + Str(RETRY1) + " -- RETRY2" + Str(RETRY2) + " -- Clipboard Logger Still Loaded Not Propering Exit -- Force Terminate 1 Min"
End Sub

Public Sub Timer1_Timer()

    RETRY1 = RETRY1 + 1
    RETRY2 = RETRY2 + 1
    RETRY3 = RETRY3 + 1
    Label1.Caption = "RETRY " + Str(RETRY1) + " -- RETRY2" + Str(RETRY2) + " -- Clipboard Logger Still Loaded Not Propering Exit -- Force Terminate 1 Min"
    'Stop
    'Exit Sub
    EXIT_TRUE = True
    
    'End

    Unload Me
    Unload Form1
    

End Sub

Private Sub Timer2_Timer()

    Unload Me
    Unload Form1
End Sub
