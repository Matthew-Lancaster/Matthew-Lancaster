Attribute VB_Name = "Shell32Win"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Public vFileName As String

' Eample of using ShellExecute
' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal

Public Sub vLaunch(vFile As String)
ShellExecute hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal
End Sub
