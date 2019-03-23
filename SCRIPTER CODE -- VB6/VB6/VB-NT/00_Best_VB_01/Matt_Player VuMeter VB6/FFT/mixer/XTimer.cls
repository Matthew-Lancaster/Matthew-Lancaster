VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class contains all of the timer propertys and methods
' Author: Matt Gillore (SCO_STINKS@Yahoo.com)
Option Explicit
Private mlngTimerID As Long
Private mlngInterval As Long
Private mblnEnabled As Boolean
Private mTag

Event Timer()


Friend Property Get TimerID() As Long
    TimerID = mlngTimerID
End Property

Public Property Get Tag() As Variant
    Tag = mTag
End Property

Public Property Let Tag(nTag As Variant)
    mTag = nTag
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = 0
    Enabled = mblnEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    If NewValue = mblnEnabled Then Exit Property
    mblnEnabled = NewValue
    If mlngInterval = 0 Then Exit Property

    If mblnEnabled Then
        Debug.Assert mlngTimerID = 0
        mlngTimerID = BeginTimer(Me, mlngInterval)
    Else
        If mlngTimerID <> 0 Then
            Call EndTimer(Me)
            mlngTimerID = 0
        End If
    End If
End Property

Public Property Get Interval() As Long
    Interval = mlngInterval
End Property
'
Public Property Let Interval(ByVal NewInterval As Long)
    If NewInterval = mlngInterval Then Exit Property
    mlngInterval = NewInterval
    
    If mlngTimerID <> 0 Then
        Call EndTimer(Me)
        mlngTimerID = 0
    End If
    
    If (NewInterval <> 0) And mblnEnabled Then
        mlngTimerID = BeginTimer(Me, NewInterval)
    End If
End Property

Public Sub RaiseTick()
    RaiseEvent Timer
End Sub



Private Sub Class_Terminate()

    On Error Resume Next
    If mlngTimerID <> 0 Then KillTimer 0, mlngTimerID

End Sub

