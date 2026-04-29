VERSION 5.00
Begin VB.Form Form_RUN_GOODSYNC 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   14616
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   14616
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "ELITE SPY MAINTAIN ____ RUNNER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15336
   End
End
Attribute VB_Name = "Form_RUN_GOODSYNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COUNT_DOWN
Dim ENDER_HAPPEN


Private Sub Form_Load()
COUNT_DOWN = 40
Call Timer1_Timer
Me.Show

XW = Me.Height
YH = Me.Width
Debug.Print XW '4300
Debug.Print YH '15000
XW = 4300
YH = 15000
Me.Height = XW
Me.Width = YH


frmMain.WindowState = vbMinimized

ReDim Preserve BLOCK_RUN_1(20)
ReDim Preserve BLOCK_RUN_2(20)
ReDim Preserve BLOCK_RUN_3(20)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If ENDER_HAPPEN = True Then Exit Sub

VAR_ARRAY = VAR_ARRAY + 1
BLOCK_RUN_1(VAR_ARRAY) = Now + TimeSerial(0, 1, 0)
BLOCK_RUN_2(VAR_ARRAY) = APP_NAME_RELOAD_IT_ER_EXE
BLOCK_RUN_3(VAR_ARRAY) = False
Timer1.Enabled = False
Unload Me

End Sub

Private Sub Label1_Click()

Timer1.Enabled = False
Unload Me
End

End Sub

Private Sub Label1_DblClick()
VAR_ARRAY = VAR_ARRAY + 1
BLOCK_RUN_1(VAR_ARRAY) = Now + TimeSerial(1, 0, 0)
BLOCK_RUN_2(VAR_ARRAY) = APP_NAME_RELOAD_IT_ER_EXE
BLOCK_RUN_3(VAR_ARRAY) = False
Timer1.Enabled = False
Unload Me

End Sub

Private Sub Timer1_Timer()

COUNT_DOWN = COUNT_DOWN - 1


APP_NAME = "GOODSYNC DESKTOP ____"
APP_NAME = "PROCESS LASSO ____"

APP_NAME_RELOAD_IT_ER____ = APP_NAME_RELOAD_IT_ER____
APP_NAME_RELOAD_IT_ER_EXE = APP_NAME_RELOAD_IT_ER_EXE

Label1.Caption = Trim(Str(COUNT_DOWN)) + vbCrLf + " ELITE SPY __ MAINTAIN DO_ER" + vbCrLf + APP_NAME_RELOAD_IT_ER____ + vbCrLf + "IS RUNNER" + vbCrLf + "HITT BUTTON HERE TO EXIT" + vbCrLf + "DOUBLE BUTTON TO TIMER" + vbCrLf + "AND CLOSE TO SHORT TIMER"
Me.Caption = Replace(Replace(Label1.Caption, vbCrLf, " "), "__", "_")

If COUNT_DOWN > 0 Then Exit Sub

'APP_NAME_RELOAD_IT_ER_EXE = "C:\Program Files\Process Lasso\ProcessLasso.exe"

Shell APP_NAME_RELOAD_IT_ER_EXE, vbMinimizedNoFocus

Timer1.Enabled = False
ENDER_HAPPEN = True
Unload Me

End Sub



