VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   960
      Top             =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Gabo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim i As Double
Dim cnt As Integer
Dim ra As Integer
Dim x, y As Integer
Dim h As Double
Dim j As Double
Dim k As Double

Private Declare Function ShowCursor Lib "User32" (ByVal Show As Integer) As Integer
Private Sub Form_Load()
    
    ShowCursor 0
    Randomize
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
    i = 0
    cnt = 1
    x = Rnd * frmMain.Width
    y = Rnd * frmMain.Height
    


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
'    Static Count As Integer: Count = Count + 1: If Count > 2 Then End
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
'    Static Count As Integer: Count = Count + 1: If Count > 2 Then End
 
End Sub

Private Sub Timer4_Timer()

On Error GoTo err
    
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
    i = 0
    h = 12000
    j = 0
    k = 9000
    cnt = 0
    Do
        DoEvents
        Me.Line (i, 0)-(i, 9000), RGB(R, G, B)
        Me.Line (h, 0)-(h, 9000), RGB(R, G, B)
        Me.Line (0, j)-(12000, j), RGB(R, G, B)
        Me.Line (0, k)-(12000, k), RGB(R, G, B)
        i = i + 1.33
        h = h - 1.33
        j = j + 1
        k = k - 1
        cnt = cnt + 1
        If cnt >= 20 Then
            R = R - 2
            G = G - 2
            B = B - 2
            cnt = 0
        End If
    Loop While i < 6000 And h > 6000 And j < 4500 And k > 4500
    i = 0
Exit Sub

err:
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
    Resume Next

End Sub

