VERSION 5.00
Begin VB.Form frmDX 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form2"
   ScaleHeight     =   4365
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDX
' DateTime  : 4/2/2005 12:26
' Author    : Jason
' Purpose   : A plain form for DX handle and mouse events
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyEscape Then
      bRunning = False
      EndIt
   End If
   
   If KeyCode = vbKeyF Then
      bFPS = Not (bFPS)
   End If
      
End Sub

Private Sub Form_Load()

   Me.Top = 0
   Me.Left = 0
   Me.Width = DX_WIDTH * Screen.TwipsPerPixelX
   Me.Height = DX_HEIGHT * Screen.TwipsPerPixelX
   Me.ScaleMode = vbPixels
   KeyPreview = True
   ShowCursor 0
   
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   DrawingMenu = False
   
   If DrawingMeter = False Then
      PaddleReceive = True
   End If
   
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


   PaddleA.X = X
   If PaddleA.X + PADDLE_WIDTH > BACK_WIDTH + BACK_OFFSET Then
      PaddleA.X = BACK_WIDTH + BACK_OFFSET - PADDLE_WIDTH
   End If
   
   If PaddleA.X < 0 + BACK_OFFSET Then
      PaddleA.X = BACK_OFFSET
   End If
   
End Sub

Public Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   DrawingMenu = False
   
'   BallStick = True
   If BallStick = True Then
      BallStick = False
      FireStraightUp = True
      PaddleReceive = False
      DrawingMeter = True
      PlayPaddleSound
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   ShowCursor 1
   
End Sub
