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




' ----------------------------------------------------------------------------
' excel - How to insert a break line in 'mshta.exe' pop up (VBA macro)? - Stack Overflow 
' https://stackoverflow.com/questions/54253640/how-to-insert-a-break-line-in-mshta-exe-pop-up-vba-macro/57512186#57512186
' ----------------------------------------------------------------------------
' NOTE THESE TWO LINE DIFFER
' ----------------------------------------------------------------------------
' Private Declare PtrSafe Function MsgBoxTimeout _
' ----------------------------------------------------------------------------
' Private Declare Function MsgBoxTimeout _
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





----
Blackview A60 Mobile Phone Android Smartphone, Dual SIM Free Android Phone Unlocked, 6.1-Inch IPS Full-Screen, 4080mAh Battery, 5MP+13MP Dual Camera, Android 8.1, UK Version - Black: Amazon.co.uk: Electronics 
https://www.amazon.co.uk/Blackview-SIM-Free-Smartphones-Unlocked-Waterdrop/dp/B07QGJXJLN/ref=bmx_6/257-2107926-6902857?_encoding=UTF8&pd_rd_i=B07QGJXJLN&pd_rd_r=908232eb-4155-426e-b96e-34f62370c9d8&pd_rd_w=16xfL&pd_rd_wg=UNnJn&pf_rd_p=26d9b5cc-9c88-4655-88a4-db5806d5daad&pf_rd_r=FS1B1KEDH9DZNQRM0HKV&refRID=FS1B1KEDH9DZNQRM0HKV&th=1
----
Lee
5.0 out of 5 stars Blackview A60 - Excellent budget mobile phone
Reviewed in the United Kingdom on 18 February 2020
Colour Name: BlackVerified Purchase
Awesome phone for the price. Android Go OS very familiar to use. Go doesn't seem to be much different to the full version in day to day use. For phone calls and light internet use it's ideal. Dual SIM is what drew me to it, one for business one for pleasure. Overall features and quality are very good and comparable with much more expensive phones. Battery is larger capacity than similar models which is always a good thing. Silicon cover and screen protector included so that's another bonus. You simply can't beat it for a budget phone! *****

Handy Hint! You can speed the phone up even more if you wish...

1. Go to SYSTEM
2. Go to BUILD.
3. Press Build 7 times.
4. Now you have DEVELOPER options.
5. Go back one level.
6. Go to DEVELOPER options.
7. Scroll down to the 'animation' settings (there are 3).
8. Turn the 3 animation settings from 1x to OFF.
9. Press back out of settings and use your faster phone!

Note. Some people advise not to turn animation all the way off but turn it to 0.5x Try it yourself and see what's best for you. I've had no problems with it off and the phone is lightning quick. Mess with other developer options at your own risk. It may also void your warranty but you can return everything back to normal with a factory reset. You can use this trick on most things Android.

Well done Blackview! An excellent phone at an awesome price!
Read less
55 people found this helpful
