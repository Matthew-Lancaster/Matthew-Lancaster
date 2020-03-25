VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Declare Function ClipCursor Lib "user32" (lpRect As Any) _
        As Long
        
Private Declare Function GetClipCursor Lib "user32" (lpRect As RECT) _
        As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As _
        Long, lpRect As RECT) As Long
       
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private m_Left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long

Public Property Get Left() As Long
    Call GetClipRect
    Left = m_Left
End Property

Public Property Let Left(ByRef Value As Long)
    m_Left = Value
    Call SetClipRect
End Property

Public Property Get Top() As Long
    Call GetClipRect
    Top = m_Top
End Property

Public Property Let Top(ByRef Value As Long)
    m_Top = Value
    Call SetClipRect
End Property

Public Property Get Width() As Long
    Call GetClipRect
    Width = m_Width
End Property

Public Property Let Width(ByRef Value As Long)
    m_Width = Value
    Call SetClipRect
End Property

Public Property Get Height() As Long
    Call GetClipRect
    Height = m_Height
End Property

Public Property Let Height(ByRef Value As Long)
    m_Height = Value
    Call SetClipRect
End Property

Public Sub Release()
    Call ClipCursor(ByVal 0&)
End Sub

Public Sub Control(ByRef hwnd As Long)
  Dim R As RECT, P As POINTAPI

    Call GetWindowRect(hwnd, R)
    Call ClipCursor(R)
End Sub

Private Sub SetClipRect()
    Dim R As RECT
    
        R.x1 = m_Left
        R.y1 = m_Top
        R.x2 = m_Width - m_Left
        R.y2 = m_Height - m_Top
        
        Call ClipCursor(R)
End Sub

Private Sub GetClipRect()
    Dim R As RECT
    
        Call GetClipCursor(R)
    
        m_Left = R.x1
        m_Top = R.y1
        m_Width = R.x2 - R.x1
        m_Height = R.y2 - R.y1
End Sub
