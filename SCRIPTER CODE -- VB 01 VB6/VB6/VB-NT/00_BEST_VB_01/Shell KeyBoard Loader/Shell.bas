Attribute VB_Name = "Shell32Win"
'Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const conSwNormal = 1

Public vFileName As String


' Eample of using ShellExecute
' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal

Public Sub vLaunch(vFile As String, vPara As String)
Dim RetVal As Long
Dim VhWnd, F_Ext
Dim hWndDesk As Long
Dim success As Long
  
'the desktop will be the
'default for error messages
hWndDesk = GetDesktopWindow()

    
Set FS = CreateObject("Scripting.FileSystemObject")
F_Ext = FS.GetExtensionName(vFile)


'For Each Form In Forms
'    VhWnd = Form.hwnd
'    If VhWnd > 0 Then Exit For
'Next


If LCase(F_Ext) <> "lnk" Then
    success = ShellExecute(hWndDesk, "open", vFile, vPara, vbNullString, conSwNormal)
Else
    'BOTH THESE WORK
    'success = ShellExecute(hWndDesk, vbNullString, vFile, vbNullString, "E:\", vbNormal)
    success = ShellExecute(hWndDesk, "open", vFile, vbNullString, "", vbNormal)
    
End If

'LastDLLError
If Err.LastDllError > 0 Then Debug.Print Err.LastDllError

'This is optional. Uncomment the three lines
'below to have the "Open With.." dialog appear
'when the ShellExecute API call fails
If success < 32 Then
    MsgBox "ERROR WITH OPEN SHELLEXECUTE - OKAY TO LOAD    " + vbCrLf + "Call Shell(""rundll32.exe shell32.dll,OpenAs_RunDLL "" & vFile, vbNormalFocus)"
    Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & vFile, vbNormalFocus)
End If


End Sub

