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




' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' FOR VB6
' Tue 09-Jun-2020 11:10:00
' THE CODE HERE ---- D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe
' AND BEFORE HERE -- C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Private Type ULARGE_INTEGER
  LowPart As Long
  HighPart As Long
End Type
Private Type SHQUERYRBINFO
  cbSize As Long
  i64Size As ULARGE_INTEGER
  i64NumItems As ULARGE_INTEGER
End Type
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
 
' Quick conversion chart of GB to byte
' 1 GB to byte = 1073741824 byte
' ----
' Convert GB to byte - Conversion of Measurement Units
' https://www.convertunits.com/from/GB/to/byte
' ----

Dim O_DAY_NOW_MIDNIGHT__

Private Sub Timer_1_SECOND_Timer()
	Call TIMER_O_DAY_NOW_MIDNIGHT__
END SUB

Sub TIMER_O_DAY_NOW_MIDNIGHT__()
	If O_DAY_NOW_MIDNIGHT__ = 0 Then
		O_DAY_NOW_MIDNIGHT__ = Day(Now)
	End If
	If Day(Now) = O_DAY_NOW_MIDNIGHT__ Then Exit Sub
	O_DAY_NOW_MIDNIGHT__ = Day(Now)

	If GET_RECYCLEBIN_SIZE > 0 Then
		Shell "EXPLORER shell:RecycleBinFolder", vbNormalFocus
	End If
End Sub
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------



