' ----------------------------------------------------------------------------
' Count = 430 -- Tue 09-Jun-2020 09:36:55
' ---------------------
' Form FindWindow ---
' Msgbox With Timer - VB6 | Dream.In.Code - Google Chrome
' ---------------------
' ----
' Msgbox With Timer - VB6 | Dream.In.Code 
' https://www.dreamincode.net/forums/topic/59064-msgbox-with-timer/
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' Count = 429 -- Tue 09-Jun-2020 09:36:51
' ---------------------
' Form FindWindow ---
' Msgbox With Timer - VB6 | Dream.In.Code - Google Chrome
' ----------------------------------------------------------------------------

' ----
' MsgBox constants | Microsoft Docs 
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-constants
' ----


' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' SET A MSGBOX THAT QUIT WITH TIMER
' ----------------------------------------------------------------------------
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")  ' -- VBS WAY OF PUTTER
' ----------------------------------------------------------------------------
' HERE WORK VBS BUT NOT VB6
' ----------------------------------------------------------------------------
' NEXT METHOD WORKER WITH VB6
' BUT NOT REALLY REQUIRE -- API NEXT METHOD ON
' ----------------------------------------------------------------------------
Dim RESULT
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
RESULT = WSHShell.Popup(I_1, 10, , vbOKOnly + vbMsgBoxSetForeground)
Set WSHShell = Nothing
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------



' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
ByVal hwnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal uType As Long, _
ByVal wLanguageId As Long, _
ByVal dwMilliseconds As Long _
) As Long

Private Const IDOKONLY& = 0&
Private Const IDOK& = 1&
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES& = 6&
Private Const IDNO& = 7&
Private Const MB_SETFOREGROUND& = &H10000
Private Const MB_YESNO& = &H4&
Private Const MB_ICONASTERISK& = &H40&
Private Const MB_TIMEDOUT& = &H7D00&

' ----------------------------------------------------------------------------
' SET A MSGBOX THAT QUIT WITH TIMER
' ----------------------------------------------------------------------------
Me.Show
RESULT = MessageBoxTimeout(0&, I_1, App.EXEName, MB_SETFOREGROUND Or IDOKONLY Or MB_ICONASTERISK, 0&, 10000&)
If RESULT = "IDYES" Then
End If
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------






