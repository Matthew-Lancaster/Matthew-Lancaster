VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMousePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
        
Private Declare Function SetCursorPos Lib "user32" (ByVal X _
        As Long, ByVal Y As Long) As Long
        
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private m_x As Long
Private m_y As Long

Public Property Get X() As Long
    Call GetPosition
    X = m_x
End Property

Public Property Let X(Value As Long)
    Call SetCursorPos(Value, m_y)
    m_x = Value
End Property

Public Property Get Y() As Long
    Call GetPosition
    Y = m_y
End Property

Public Property Let Y(Value As Long)
    Call SetCursorPos(m_x, Value)
    m_y = Value
End Property

Private Sub GetPosition()
    Dim P As POINTAPI
    
        Call GetCursorPos(P)
        m_y = P.Y
        m_x = P.X
End Sub
