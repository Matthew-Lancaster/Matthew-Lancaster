VERSION 5.00
Begin VB.Form Form2_TIME_SET_BACK 
   Caption         =   "Form2"
   ClientHeight    =   5340
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13404
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   13404
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TIMER_UNLOAD_FORM_SHOWING 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   10680
      Top             =   840
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2_TIME_SET_BACK.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4884
      Left            =   120
      TabIndex        =   0
      Top             =   96
      Width           =   12192
   End
End
Attribute VB_Name = "Form2_TIME_SET_BACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Label1.Left = 0
Label1.Top = 0
Label1.Width = Me.Width
Label1.Height = Me.Height

i = vbCrLf
Label1.Caption = "THE CODE TIMER TICKER EVENT CHAIN __" + i + "HAS BEEN SET BACK MORE THAN _ A MINUTE __" + i + "AND THE CLOCK WILL BE SET AGAIN IN _" + i + "A _ 1 SECOND"
'__ EDIT LABEL IN HARDCODED"

End Sub

Private Sub TIMER_UNLOAD_FORM_SHOWING_Timer()

Unload Me

'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus
Call Form1.MNU_BOOT_KILLER_Click

RUN_FROM_AUTO_TASK_KILLER = True
Call Form1.MNU_TASK_KILLER_Click



End Sub
