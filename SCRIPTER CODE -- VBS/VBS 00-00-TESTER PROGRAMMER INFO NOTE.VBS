' ---------------------
' ---------------------
' Count = 430 -- Tue 09-Jun-2020 09:36:55
' ---------------------
' Form FindWindow ---
' Msgbox With Timer - VB6 | Dream.In.Code - Google Chrome
' ---------------------
' ----
' Msgbox With Timer - VB6 | Dream.In.Code 
' https://www.dreamincode.net/forums/topic/59064-msgbox-with-timer/
' ----
' ---------------------

' ---------------------
' Count = 429 -- Tue 09-Jun-2020 09:36:51
' ---------------------
' Form FindWindow ---
' Msgbox With Timer - VB6 | Dream.In.Code - Google Chrome
' ---------------------

Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
ByVal hwnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal uType As Long, _
ByVal wLanguageId As Long, _
ByVal dwMilliseconds As Long _
) As Long

Private Const IDYES& = 6&
Private Const IDNO& = 7&
Private Const MB_SETFOREGROUND& = &H10000
Private Const MB_YESNO& = &H4&
Private Const MB_ICONASTERISK& = &H40&
Private Const MB_TIMEDOUT& = &H7D00&

' Private Sub Form_Load()
Dim lRet As Long
lRet = MessageBoxTimeout(0&, "your message", _
"Timeout MessageBox Test", MB_SETFOREGROUND Or MB_YESNO Or MB_ICONASTERISK, _
0&, 10000&)

Select Case lRet
Case IDYES:
'// your code
Case IDNO:
'// your code

Case MB_TIMEDOUT:
'// your code

End Select
End Sub

