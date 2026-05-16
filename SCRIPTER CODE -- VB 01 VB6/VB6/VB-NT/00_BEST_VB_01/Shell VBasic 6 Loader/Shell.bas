Attribute VB_Name = "Shell32Win"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Public vFileName As String

' Eample of using ShellExecute
' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal

Public Sub vLaunch(vFile As String, vPara As String)

On Error Resume Next
ShellExecute hWnd, "open", vFile, vPara, vbNullString, conSwNormal

If Err.Number > 0 Then
    MsgBox "Error Run Check With Administrator" + vbCrLf + vbCrLf + "ShellExecute hWnd, ""open"", vFile, vPara, vbNullString, conSwNormal", vbMsgBoxSetForeground
End If


End Sub
