VERSION 5.00
Begin VB.Form Form2_TIME_SET_BACK 
   Caption         =   "Form2"
   ClientHeight    =   3120
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11088
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   11088
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TIMER_UNLOAD_FORM_SHOWING 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   10680
      Top             =   840
   End
   Begin VB.Label Label1 
      Caption         =   "THE TIME HAS BEEN SET BACK MORE THAN 1 MINUTE AN WILL BE SET AGAIN IN 1 SECOND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2724
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10452
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

End Sub

Private Sub TIMER_UNLOAD_FORM_SHOWING_Timer()

Unload Me

Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus

End Sub
