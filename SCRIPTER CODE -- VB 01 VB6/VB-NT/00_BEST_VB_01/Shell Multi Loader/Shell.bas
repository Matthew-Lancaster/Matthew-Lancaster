Attribute VB_Name = "Shell32Win"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Public vFileName As String

' Eample of using ShellExecute
' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal

Public Sub vLaunch(vFile As String, vPara As String)

    
'Set fs22 = CreateObject("Scripting.FileSystemObject")
'Set F = fs22.GetFile(vFile)
'If F.Size > 1000 ^ 2 Then

'End If

'if
'C:\Program Files\TextView\Textview.exe


ShellExecute hWnd, "open", vFile, vPara, vbNullString, conSwNormal
'ShellExecute hWnd, "open", "D:\VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\NotePad instr top LF 2abbrev txt.exe", vPara, vbNullString, conSwNormal
End Sub
'D:\VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\NotePad instr top LF 2abbrev txt.exe

