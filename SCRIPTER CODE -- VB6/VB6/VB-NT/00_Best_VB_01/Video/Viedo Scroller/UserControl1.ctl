VERSION 5.00
Begin VB.UserControl Scroller 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer STimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin VB.PictureBox BackB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   3735
   End
   Begin VB.PictureBox Draw 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Scroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' SCROLLER.OCX
' Youll notice that ive had to use timers in this code, and thats
' because I couldnt find a way to create an asynchronus function
' in vb.
' You can call 21 override functions before the override will
' stop working, howvever once the 1st scroll is finished that "slot"
' will be free and override will work again.

'for example, if i only had 2 slots:
'1st override call   |<---->|<---------------------------------¬
'2nd override called    |<---->|                               /
'and if you called it a 3rd time then it would start when this
'point is reached.  (here:) |<---->|

'hope thats not too confusing, the rest of its fairly simple
'please vote for it :)
'Paul Blower
'mercior@hotmail.com

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private tScrollText As String
Private tScrollSpeed As Integer
Private tScrollPause As Long
Private tOverride As Boolean
Private Direction As Integer '0=left, 1=up

Private Scrolling As Boolean
Private CX As Integer
Private CY As Integer

Public Property Let BackColor(nBackCol As Long)
    BackB.BackColor = nBackCol
End Property

Public Property Let TextColor(nTxtCol As Long)
    BackB.ForeColor = nTxtCol
End Property

Sub PS(Interval As Long) 'Pauses for specified interval
    Start = GetTickCount
        Do While Start + Interval > GetTickCount
            DoEvents
        Loop
End Sub

Public Sub DoScroll(ScrollText As String, ScrollSpeed As Integer, ScrollPause As Long, Override As Boolean, Optional ScrollDir As Integer)
    If ScrollDir = 1 Then 'check scroll direction
        Direction = 1
    Else
        Direction = 0
    End If
    If Not Scrolling Then
        tScrollText = ScrollText
        tScrollSpeed = ScrollSpeed
        tScrollPause = ScrollPause
        tOverride = Override
    ElseIf Scrolling And Override Then
        tScrollText = ScrollText
        tScrollSpeed = ScrollSpeed
        tScrollPause = ScrollPause
        tOverride = Override
    Else
        Exit Sub
    End If
    For i = 0 To 20
        If STimer(i).Enabled = False Then
            STimer(i).Enabled = True
            Exit Sub
        End If
    Next
End Sub

Private Sub STimer_Timer(Index As Integer)
    If Direction = 1 Then
        Call MainLoopUp
    Else
        Call MainLoop
    End If
    STimer(Index).Enabled = False
End Sub

Private Sub UserControl_Initialize()
    BackB.FontSize = 8
    BackB.Font = "Fixedsys"
    BackB.CurrentX = 0
    BackB.CurrentY = 0
    BackB.BackColor = RGB(0, 0, 0)
    BackB.ForeColor = RGB(240, 240, 240)
    For i = 1 To 20
        Load STimer(i)
        STimer(i).Enabled = False
        STimer(i).Interval = 1
    Next
End Sub
Private Sub MainLoopUp()
    CX = 1
    CY = Draw.ScaleHeight
    Scrolling = True
    Do While CY > 1
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CY = CY - tScrollSpeed
        PS 20
    Loop
    
    PS tScrollPause
    
    Do While CY > -16
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CY = CY - tScrollSpeed
        PS 20
    Loop
    Scrolling = False
End Sub

Private Sub MainLoop()
    CX = Draw.ScaleWidth
    CY = 0
    Scrolling = True
    Do While CX > 1
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CX = CX - tScrollSpeed
        PS 20
    Loop
    
    PS tScrollPause
    
    Do While CX > 0 - (Len(tScrollText) * 8)
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CX = CX - tScrollSpeed
        PS 20
    Loop
    Scrolling = False
End Sub

Private Sub UserControl_Resize()
    Draw.Left = 0
    Draw.Top = 0
    Draw.Width = ScaleWidth
    Draw.Height = ScaleHeight
    BackB.Top = ScaleHeight + 1
    BackB.Width = ScaleWidth
    BackB.Height = ScaleHeight
End Sub
