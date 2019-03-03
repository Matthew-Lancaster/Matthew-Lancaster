Attribute VB_Name = "modTimer"
'---------------------------------------------------------------------------------------
' Module    : modTimer
' DateTime  : 4/2/2005 12:28
' Author    : Jason   (noi_max@hotmail.com)
' Purpose   : Module for Timing and API DoEvents
'---------------------------------------------------------------------------------------

Option Explicit

Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (LpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As Currency) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
                                      
Private Const PM_REMOVE = &H1       'paramater on peekmessage to remove or leave message in queue

'*************************

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type MSG
   hwnd     As Long        'window where message occured
   Message  As Long        'message id itself
   wParam   As Long        'further defines message
   lParam   As Long        'further defines message
   time     As Long        'time of message event
   pt       As POINTAPI    'position of mouse
End Type

Dim Message As MSG         'holds message recieved from queue


'For the timer function
Public PerfFreq As Currency
Public TimeStart As Currency
Public TimeEnd As Currency


''Variables to keep track of our timing
Dim perfStart As Currency
Dim perfEnd As Currency
Public Elapsed As Double

Public Sub QTimer()

   'Use the timer sub to get the Elapsed time between each processor tick.
   'This is measured in microseconds, so a multiplier is used
   'to get milliseconds instead
   
   'Get current time
   QueryPerformanceCounter perfStart
   
   'Calculate elapsed time
   Elapsed = (perfStart - perfEnd) / PerfFreq * 1000
   
   'Now, we start the timer again, this will be used to calculate
   'how much time has passed since we last called this subroutine.
   QueryPerformanceCounter perfEnd
    
End Sub

Public Sub myDoEvents(ByVal hwnd As Long)

   'Replacement for VB's slow DoEvents function

   If PeekMessage(Message, hwnd, 0&, 0&, PM_REMOVE) Then
   
      TranslateMessage Message
      DispatchMessage Message
   
   End If

End Sub


