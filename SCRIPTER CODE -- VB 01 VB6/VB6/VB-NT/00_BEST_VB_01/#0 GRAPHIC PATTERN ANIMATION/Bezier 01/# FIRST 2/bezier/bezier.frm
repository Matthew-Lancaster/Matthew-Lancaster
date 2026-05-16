VERSION 5.00
Begin VB.Form Bezier 
   Caption         =   "Bezierkurven"
   ClientHeight    =   5130
   ClientLeft      =   1965
   ClientTop       =   1515
   ClientWidth     =   6615
   Icon            =   "bezier.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5130
   ScaleWidth      =   6615
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Zufallskurve 
         Caption         =   "Zufallskurve"
         Height          =   345
         Left            =   60
         TabIndex        =   6
         Top             =   45
         Width           =   1290
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Copyright © by Jürgen Anke (1998)"
         Height          =   210
         Left            =   1485
         TabIndex        =   7
         Top             =   105
         Width           =   4485
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   6  'Mask Pen Not
      Height          =   4110
      Left            =   -15
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   0
      Top             =   435
      Width           =   5640
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   30
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   225
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   780
         MousePointer    =   2  'Cross
         TabIndex        =   3
         Top             =   105
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   450
         MousePointer    =   2  'Cross
         TabIndex        =   2
         Top             =   135
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   285
         MousePointer    =   2  'Cross
         TabIndex        =   1
         Top             =   405
         Width           =   120
      End
   End
End
Attribute VB_Name = "Bezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long

Dim ptBez(3) As POINTAPI
Dim Drag As Boolean
Sub Dragging(nGriff As Byte, X As Single, Y As Single)
    
    DrawBezier
    griff(nGriff).Left = X - 4
    griff(nGriff).Top = Y - 4
    ptBez(nGriff).X = X
    ptBez(nGriff).Y = Y
    DrawBezier

End Sub

Sub DrawBezier()
Dim i As Byte

    Call PolyBezier(Picture1.hdc, ptBez(0), 4)
    
    For i = 0 To 3
        griff(i).Left = ptBez(i).X - 4
        griff(i).Top = ptBez(i).Y - 4
    Next
    
    With Picture1
        .DrawStyle = vbDot
        Picture1.Line (ptBez(0).X, ptBez(0).Y)-(ptBez(1).X, ptBez(1).Y)
        Picture1.Line (ptBez(2).X, ptBez(2).Y)-(ptBez(3).X, ptBez(3).Y)
        .DrawStyle = vbSolid
    End With
    
End Sub


Sub RandomPoints()
Dim MaxX As Long, MaxY As Long

    Randomize
    MaxX = Picture1.ScaleWidth
    MaxY = Picture1.ScaleHeight
    ptBez(0).X = Int(Rnd * MaxX)
    ptBez(0).Y = Int(Rnd * MaxY)
    ptBez(1).X = Int(Rnd * MaxX)
    ptBez(1).Y = Int(Rnd * MaxY)
    ptBez(2).X = Int(Rnd * MaxX)
    ptBez(2).Y = Int(Rnd * MaxY)
    ptBez(3).X = 300
    ptBez(3).Y = 140
    
End Sub



Private Sub Zufallskurve_Click()

    Picture1.Cls
    RandomPoints
    DrawBezier
    
End Sub
Private Sub Form_Load()
    
    RandomPoints
    
End Sub


Private Sub Form_Resize()
Picture1.Move Picture1.Left, Picture1.Top, ScaleWidth, ScaleHeight - Picture2.Height
End Sub


Private Sub griff_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    griff(Index).Drag
    Drag = True
End If
End Sub


Private Sub griff_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Drag = False
End Sub


Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    
    Call Dragging(Source.Index, X, Y)
    
End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    Call Dragging(Source.Index, X, Y)
    
End Sub

Private Sub Picture1_Paint()

    If Not Drag Then DrawBezier
    
End Sub


