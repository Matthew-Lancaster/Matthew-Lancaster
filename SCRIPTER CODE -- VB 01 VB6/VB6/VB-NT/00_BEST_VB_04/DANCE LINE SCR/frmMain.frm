VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDance 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin VB.Line linLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   480
      X2              =   3840
      Y1              =   1800
      Y2              =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Sub tmrDance_Timer()
    Static DirX1 As Integer, Speedx1 As Integer
    Static DirX2 As Integer, Speedx2 As Integer
    Static DirY1 As Integer, Speedy1 As Integer
    Static DirY2 As Integer, Speedy2 As Integer
    
    'Set initial Direction
    If DirX1 = 0 Then DirX1 = Rnd * 3
    If DirX2 = 0 Then DirX2 = Rnd * 3
    If DirY1 = 0 Then DirY1 = Rnd * 3
    If DirY2 = 0 Then DirY2 = Rnd * 3
    
    'Set Speed
    If Speedx1 = 0 Then Speedx1 = 60 * Int(Rnd * 5)
    If Speedx2 = 0 Then Speedx2 = 60 * Int((Rnd * 5))
    If Speedy1 = 0 Then Speedy1 = 60 * Int((Rnd * 5))
    If Speedy2 = 0 Then Speedy2 = 60 * Int((Rnd * 5))
    
    'Handle Movement
    'If X1=1 then moving right else moving left
    'If X2=1 then moving right else moving left
    'If Y1=1 then moving down else moving up
    'If Y2=1 then moving down else moving up
    If DirX1 = 1 Then
        linLine.X1 = linLine.X1 + Speedx1
    Else
        linLine.X1 = linLine.X1 - Speedx1
    End If
    If DirX2 = 1 Then
        linLine.X2 = linLine.X2 + Speedx2
    Else
        linLine.X2 = linLine.X2 - Speedx1
    End If
    If DirY1 = 1 Then
        linLine.Y1 = linLine.Y1 + Speedy1
    Else
        linLine.Y1 = linLine.Y1 - Speedy1
    End If
    If DirY2 = 1 Then
        linLine.Y2 = linLine.Y2 + Speedy2
    Else
        linLine.Y2 = linLine.Y2 - Speedy2
    End If
    
    'Handle bouncing (change directions if off screen)
    If linLine.X1 < 0 Then DirX1 = 1
    If linLine.X1 > Me.ScaleWidth Then DirX1 = 2
    If linLine.X2 < 0 Then DirX2 = 1
    If linLine.X2 > Me.ScaleWidth Then DirX2 = 2
    If linLine.Y1 < 0 Then DirY1 = 1
    If linLine.Y1 > Me.ScaleHeight Then DirY1 = 2
    If linLine.Y2 < 0 Then DirY2 = 1
    If linLine.Y2 > Me.ScaleHeight Then DirY2 = 2

End Sub
Sub AlwaysOnTop(FrmID As Form, OnTop As Boolean)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

If OnTop Then
OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
Else
OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End If
End Sub



Private Sub Form_Load()
    'Get current settings
    LineColor = GetSetting("Dancing Line Screen Saver", "Settings", "LINE_COLOR", 14)
    LineThickness = GetSetting("Dancing Line Screen Saver", "Settings", "LINE_THICKNESS", 1)
    LineSpeed = GetSetting("Dancing Line Screen Saver", "Settings", "LINE_SPEED", 50)
    
    'Check command line parameters
    Select Case LCase(Left(Command, 2))
        Case "/p"
            Me.Hide
            tmrDance.Interval = 0
            End
            Exit Sub
        Case "/s"
            'Proceed
        Case Else
            'Show settings
            Me.Hide
            frmSettings.Show
            tmrDance.Interval = 0
            Exit Sub
    End Select

    Dim X As Integer
    
    'Make the Screen Saver topmost
    'Call AlwaysOnTop(Me, True)
    
    'Hide the mouse cursor
    'X = ShowCursor(False)
    
    'Apply current setting to the Screen Saver
    linLine.BorderColor = QBColor(LineColor)
    linLine.BorderWidth = LineThickness
    tmrDance.Interval = 1000 / LineSpeed

    
    Call frmSettings.cmdOK_Click

    Me.Show

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    Static Count As Integer
    
    'If the mouse is moved, unload the Screen Saver
    
    'Give enough time for program to run
    Count = Count + 1
    If Count > 5 Then
        'Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If a key is pressed, unload the Screen Saver
    'Unload Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim X As Integer
    
    'Show the mouse cursor again
    X = ShowCursor(True)
    
    'End the program
    End

End Sub


