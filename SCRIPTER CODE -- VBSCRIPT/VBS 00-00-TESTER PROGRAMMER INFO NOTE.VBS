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
RESULT = MessageBoxTimeout(0&, MSGBOX_STRING, App.EXEName+ ".EXE", MB_SETFOREGROUND Or IDOKONLY Or MB_ICONASTERISK, 0&, 10000&)
If RESULT <> IDOK Then
	Exit Sub
End If
RESULT = MessageBoxTimeout(0&, MSGBOX_STRING, App.EXEName + ".EXE", MB_SETFOREGROUND Or MB_YESNO Or MB_ICONASTERISK, 0&, 10000&)
If RESULT <> IDYES Then
	Exit Sub
End If
If RESULT <> IDYES Then
	Exit Sub
End If
If RESULT = IDNO Then
	Exit Sub
End If
If RESULT = IDCANCEL Then
	Exit Sub
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





' This part needs to be at the top of a VBA module
#If Win64 Then 
    Private Declare PtrSafe Function MsgBoxTimeout _
        Lib "user32" _
        Alias "MessageBoxTimeoutA" ( _
            ByVal hwnd As LongPtr, _
            ByVal lpText As String, _
            ByVal lpCaption As String, _
            ByVal wType As VbMsgBoxStyle, _
            ByVal wlange As Long, _
            ByVal dwTimeout As Long) _
    As Long
#Else
    Private Declare Function MsgBoxTimeout _
        Lib "user32" _
        Alias "MessageBoxTimeoutA" ( _
            ByVal hwnd As Long, _
            ByVal lpText As String, _
            ByVal lpCaption As String, _
            ByVal wType As VbMsgBoxStyle, _
            ByVal wlange As Long, _
            ByVal dwTimeout As Long) _
    As Long
#End If


Sub TestMsgbox()
    Dim ReturnValue

    ReturnValue = MsgBoxTimeout(0, "Do you like this message?" & vbCrLf & "This message box will be closed after 4 seconds." & vbCrLf & vbCrLf & "(See Immediate window for return value)", "Return Choice", vbQuestion + vbYesNoCancel, 0, 4000)
    Select Case ReturnValue
        Case vbYes
            Debug.Print "You picked Yes."
        Case vbNo
            Debug.Print "You picked No."
        Case vbCancel
            Debug.Print "You picked Cancel."
        Case 32000
            Debug.Print "Timeout before user made selection."
    End Select
End Sub
