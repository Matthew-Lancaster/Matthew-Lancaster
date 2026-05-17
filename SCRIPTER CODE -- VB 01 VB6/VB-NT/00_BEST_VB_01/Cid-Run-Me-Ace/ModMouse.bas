Attribute VB_Name = "modMouse"


Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As _
        Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As _
        Long, ByVal nIDEvent As Long) As Long

Public TimerEnabled As Boolean

Private colObj As New Collection
Private hTimer As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hCur As Long) As Long
Public hOldSysCursor2 As Long
Public jade As Integer


Private Sub Init()
    If hTimer = 0 Then
        hTimer = SetTimer(0, 0, 20&, AddressOf TimerProc)
        TimerEnabled = True
    End If
End Sub

Private Sub Terminate()
    Call KillTimer(0, hTimer)
    TimerEnabled = False
    hTimer = 0
End Sub




Private Sub TimerProc(ByVal hWnd As Long, ByVal msg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Static Flag As Boolean
    
        If Not Flag Then
            Flag = True
            Call MakeEvents
            Flag = False
        End If
End Sub

Public Function AddObject(ByRef Obj As clsMouse, Key As String) As Boolean
    Static Counter As Long
    
        Call Init
        Counter = Counter + 1
        Key = "x" & CStr(Counter)
        colObj.Add Obj, Key
        AddObject = True
End Function

Public Function RemoveObject(ByVal Key As String) As Boolean
    Dim Obj As clsMouse, z As Long
    
        colObj.Remove Key
        
        For Each Obj In colObj
            z = z + 1
        Next Obj
        If z = 0 Then Call Terminate
End Function

Public Sub MakeEvents()
    Dim Obj As clsMouse, z As Long

        For Each Obj In colObj
            Obj.TimerEvent = True
            z = z + 1
        Next Obj
        If z = 0 Then Call Terminate
End Sub
