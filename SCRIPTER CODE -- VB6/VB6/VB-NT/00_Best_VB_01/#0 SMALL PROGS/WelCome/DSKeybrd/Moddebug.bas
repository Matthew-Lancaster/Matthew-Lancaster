Attribute VB_Name = "modDebug"
'*********************************************************
Option Explicit
'*********************************************************
Private Declare Sub OutputDebugString Lib "kernel32" Alias _
              "OutputDebugStringA" (ByVal lpOutputString$)
'*********************************************************
Sub DPrint(ParamArray v())
Static dummy$, i&

dummy = ""
If Not IsMissing(v) Then
  For i = 0 To UBound(v)
    dummy = dummy & vbTab & CStr(v(i))
  Next
End If
OutputDebugString dummy
End Sub

'*********************************************************

