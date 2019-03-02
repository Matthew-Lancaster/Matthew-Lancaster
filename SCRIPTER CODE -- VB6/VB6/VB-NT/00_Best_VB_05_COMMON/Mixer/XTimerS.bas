Attribute VB_Name = "XTimerSupport"
'Simple timer implementation
'Author: Matt Gillmore (SCO_STINKS@Yahoo.com)


Option Explicit

Const MAXTIMERINCREMEMT = 10

Private Type XTIMERINFO
    xt As XTimer
    id As Long
    blnReentered As Boolean
End Type

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private maxti() As XTIMERINFO
Private mintMaxTimers As Integer

Public Function BeginTimer(ByVal xt As XTimer, ByVal Interval As Long)
    Dim lngTimerID As Long
    Dim intTimerNumber As Integer
    
    lngTimerID = SetTimer(0, 0, Interval, AddressOf TimerProc)
    If lngTimerID = 0 Then Err.Raise vbObjectError + 31013, , "Keine Zeitgeber verfügbar."
    For intTimerNumber = 1 To mintMaxTimers
        If maxti(intTimerNumber).id = 0 Then Exit For
    Next
    If intTimerNumber > mintMaxTimers Then
        mintMaxTimers = mintMaxTimers + MAXTIMERINCREMEMT
        ReDim Preserve maxti(1 To mintMaxTimers)
    End If
    Set maxti(intTimerNumber).xt = xt
    maxti(intTimerNumber).id = lngTimerID
    maxti(intTimerNumber).blnReentered = False
    BeginTimer = lngTimerID
End Function

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Dim intCt As Integer

    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = idEvent Then
            If maxti(intCt).blnReentered Then Exit Sub
            maxti(intCt).blnReentered = True
            On Error Resume Next
            maxti(intCt).xt.RaiseTick
            If Err.Number <> 0 Then
                KillTimer 0, idEvent
                maxti(intCt).id = 0
                Set maxti(intCt).xt = Nothing
            End If
            maxti(intCt).blnReentered = False
            Exit Sub
        End If
    Next
    KillTimer 0, idEvent
End Sub

Public Sub EndTimer(ByVal xt As XTimer)
    Dim lngTimerID As Long
    Dim intCt As Integer

    lngTimerID = xt.TimerID
    If lngTimerID = 0 Then Exit Sub
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = lngTimerID Then
            KillTimer 0, lngTimerID
            Set maxti(intCt).xt = Nothing
            maxti(intCt).id = 0
            Exit Sub
        End If
    Next
End Sub

Public Sub Scrub()
    Dim intCt As Integer
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id <> 0 Then KillTimer 0, maxti(intCt).id
    Next
End Sub
