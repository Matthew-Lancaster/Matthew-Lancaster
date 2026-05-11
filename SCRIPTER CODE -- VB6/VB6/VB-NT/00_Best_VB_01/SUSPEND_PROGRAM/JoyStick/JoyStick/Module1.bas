Attribute VB_Name = "JOYMOD"
Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

Public JCODE

Public Type JOYINFO
   X As Long
   Y As Long
   Z As Long
   Buttons As Long
End Type
Public Type JOYCAPS
   wMid As Integer
   wPid As Integer
   szPname As String * 32
   wXmin As Long
   wXmax As Long
   wYmin As Long
   wYmax As Long
   wZmin As Long
   wZmax As Long
   wNumButtons As Long
   wPeriodMin As Long
   wPeriodMax As Long
 End Type

Public Declare Function joyGetDevCapsA Lib "winmm.dll" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Declare Function joyGetDevCapsW Lib "winmm.dll" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type B
B(8) As Boolean
End Type


Public Function GetBPress(Num As Long, Buttons As B)
Dim s As Long
For n = 1 To 8
Buttons.B(n) = False
Next
If Num = 0 Then Exit Function
If Num = 1 Then
Buttons.B(1) = True
Exit Function
End If
If Num = 2 Then
Buttons.B(2) = True
Exit Function
End If

Do
    n = 1
    s = Num \ 2
    Do While s >= 2
    n = n + 1
    s = s \ 2
    Loop

    k = 2 ^ n
    Select Case k
    Case 1
    Buttons.B(1) = True
    Case 2
    Buttons.B(2) = True
    Case 4
    Buttons.B(3) = True
    Case 8
    Buttons.B(4) = True
    Case 16
    Buttons.B(5) = True
    Case 32
    Buttons.B(6) = True
    Case 64
    Buttons.B(7) = True
    Case 128
    Buttons.B(8) = True
    End Select

    P = Num - k
    Num = P
    
If P = 1 Then
    Buttons.B(1) = True
Exit Do
ElseIf P = 0 Then
Exit Do
End If
Loop
End Function
Public Function IsWinNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    IsWinNT = (OSVer.dwPlatformId = 2)
    
End Function
