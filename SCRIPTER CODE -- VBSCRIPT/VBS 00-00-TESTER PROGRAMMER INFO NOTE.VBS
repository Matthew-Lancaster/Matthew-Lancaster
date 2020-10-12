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

WSCRIPT.QUIT



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
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' -------------------------------------------------------------------
' HERE ARE PERFECT CODE
' NOT FOUND ON-LINE EVER EXCEPT THIS ONE
' YOU ABLE CALL ROUTINE FROM OUT ROUTINE
' OR GIVE FOLDER WITHIN
' THERE NONE NEST BESIDE IT LOOP ITSELF
' NONE MULTIPLE ROUTINE EFFORT
' -------------------------------------------------------------------
' EASY PASS 1ST PARAMETER -- WHAT FOLDER
' -------------------------------------------------------------------
' IF YOU AIN'T KICKING TOP ROUTINE SCANNER 
' LOOK AT MY SHIT
' -------------------------------------------------------------------
' WITH AR SOME EXAMPLE WHAT TO DO WITH INNER
' -------------------------------------------------------------------
' [ WED 12:44:10 PM_07 OCTOBER 2020 ]
' -------------------------------------------------------------------
' -------------------------------------------------------------------
DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")

DIM GETUSERNAME
DIM GETCOMPUTERNAME
DIM COMPUTERNAME
DIM WSHSHELL
SET WSHSHELL = CreateObject("WScript.Shell")
GETUSERNAME = CreateObject("WScript.Network").USERNAME
GETCOMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
COMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
' ------------------------------------------------------------------
' ANOTHER EXAMPLE
' ------------------------------------------------------------------
' COMPUTERNAME = WSHSHELL.ExpandEnvironmentStrings("%computername%")
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' EXAMPLE 0001 CALL ROUTINE WITH PATH
' ------------------------------------------------------------------
FOLDER_NAME="C:\Users\"+GETUSERNAME+"\AppData\Local\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\LocalState\DeviceSearchCache"

' ------------------------------------------------------------------
' EXAMPLE 0002 CALL ROUTINE NOTHING PATH AND THEN MINIMAL ONE SHOT SET PATH WITHIN 
' ------------------------------------------------------------------
FOLDER_NAME=""

COUNTER_RR=0
COUNTER_TOTAL=0

CALL SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER_NAME)

MSGBOX "DELETE DONE "& VBCR & COUNTER_RR & VBCR & "OF" & VBCR & COUNTER_TOTAL

WSCRIPT.QUIT

SUB SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)

	DIM SUBFOLDER
	DIM FILE
	ON ERROR RESUME NEXT
	
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT -- FOLDER OBJECT 
		' ABLE ARGUMENT PARAMETER AS STRING ANYTIME AND EXTRA CONVERT TAKE PLACE
		' USUAL ONCE AS STRING AND FOREVER AFTER FOLDER OBJECT VARIABLE TYPE
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		' FOLDER= "LET ME FIND A PATH"
		IF FOLDER="" THEN 
			FOLDER="E:\HARDWARE"
		END IF
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)
		EXIT SUB
	END IF
	' ---------------------------------------------------------------

	FOR EACH FILE IN FOLDER.FILES
		' DO STUFF
		' ----------------------------------------------
		' NOT GO IN DIRECTORY THAT DON'T HAVE ACCESS
		' LIKE \SYSTEM INFORMATION FOLDER FROM ROOT
		' OR BOMB OUT OF SCAN
		' SAFE NOT PROCESS ANY INTRICATE INFO ABOUT THEM
		' ----------------------------------------------
		
		FILE_PATH_STRING=FILE.path
		STR_FIND="-- .JPG .MP4" 
		STR_FIND="-- .TXT" 
		PATH_FIND_01="\DCIM"
		PATH_FIND_02="\MP_ROOT"
		EXTENSION_PART=MID(FILE.path,INSTRREV(FILE.path,"."))
		PATH_FIND_RESULT=FALSE
		
		IF INSTR(STR_FIND , UCASE(EXTENSION_PART)) >0 THEN PATH_FIND_RESULT=TRUE
		' IF INSTR(UCASE(FILE.path) , PATH_FIND_01) >0 THEN PATH_FIND_RESULT=TRUE
		' IF INSTR(UCASE(FILE.path) , PATH_FIND_02) >0 THEN PATH_FIND_RESULT=TRUE
		
		IF PATH_FIND_RESULT=TRUE THEN
			ERR.CLEAR
			FILEMODDATE=0
			SET F = FSO.GETFILE(FILE.path)
				FILEMODDATE = F.DATELASTMODIFIED
				FILECREATEDATE = F.DATECREATED
			SET F=NOTHING
			' --------------------------------------------
			' EXAMPLE DELETE FILE OVER 48 HOUR IN THE PAST
			' --------------------------------------------
			' IF FILEMODDATE+TIMESERIAL(48,0,0)<NOW THEN 
			' --------------------------------------------
			' EXAMPLE DELETE FILE OVER 20 DAY IN THE PAST
			' --------------------------------------------
			IF FILEMODDATE+20<NOW THEN 
				ON ERROR RESUME NEXT
				ERR.CLEAR
				MSGBOX FILE.path
				'''' FSO.DeleteFile FILE.path
				COUNTER_RR=COUNTER_RR+1
			ELSE
			END IF
		END IF
		COUNTER_TOTAL=COUNTER_TOTAL+1
	NEXT

	' ---------------------------------------------------------------
	' CALL GET_DRIVE_SPACE(FOLDER.path)
	' ---------------------------------------------------------------
	
	' ---------------------------------------------------------------
	' DO STUFF ON FOLDER HERE
	' ---------------------------------------------------------------
	MYFOLDER = FOLDER.path
	' ---------------------------------------------------------------
	IF MYFOLDER = "EXCELLENT" THEN 
		' DO STUFF
	END IF
	' ---------------------------------------------------------------
	
	' ---------------------------------------------------------------
    ' RECURSE INTO AR SUB FOLDER COLLECTION
	' ---------------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(SUBFOLDER)
    NEXT

END SUB

' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
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



