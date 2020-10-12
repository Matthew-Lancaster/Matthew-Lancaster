Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

' -------------------------------------------------------------------
' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF ARE
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' NEW SPAWN OF FILE 
' VBS 08-GOODSYNC EDITOR_4 MAKE SCRIPT OF JOB.VBS

' WORK TIME
' Sun 28-Jun-2020 04:18:40
' Sun 28-Jun-2020 06:49:04 -- 2 HOUR 30 MINUTE
' -------------------------------------------------------------------


' -------------------------------------------------------------------
DIM GETUSERNAME
DIM ComputerName
DIM objShell
DIM objWMIService
DIM colProcesses
DIM i1
DIM i2
DIM objProcess
DIM VBV
VBV=VBCRLF
DIM strScriptName
DIM PID_Script

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

' -------------------------------------------------------------------
' 01 OF 02 
' SOME CONFLICT WAR IN CODE STYLE 
' SOME PEOPLE TALK THIS HAS ERROR 
' THIS ONE IS CORRECT 
' -------------------------------------------------------------------
On Error Resume Next
Dim iMyPID : iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
On Error Goto 0

' -------------------------------------------------------------------
' GOOD CODE TO PASTE ON VB6
' -------------------------------------------------------------------
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
        strScriptName = Left(strScriptName, Len(strScriptName) - 1)
		PID_Script = objProcess.ProcessID
		
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		END IF
		IF strScriptName <> Wscript.ScriptName THEN
			i1=i1+1
		END IF
		
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
			' Wscript.Quit
		END IF
		IF i2 > 0 THEN ' MY  PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			' IF lngMyProcessId<>iMyPID THEN 
				' I= "ERROR MISMATCH TRUE MYPID RESULT "
				' MSGBOX I +VBV+ CSTR(iMyPID) +VBV+ CSTR(lngMyProcessId)
			' -------------------------------------------------------
			IF iMyPID<>PID_Script THEN 
				objProcess.Terminate() 
				objShell.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
				' MSGBOX CSTR(iMyPID) +VBV+ strScriptName +VBV+ CSTR(PID_Script)  +VBV+ Wscript.ScriptName
			END IF
		END IF
    END IF
NEXT



'MsgBox "Done Total " & CHR(&H64) & " Proecesses"


' ---------------------------------------------------------------------
' WORK TODAY
' ADD CODE FOR SELECTIVE FILTER _ ONLY FILTER SOME JOB FOR ACTION TAKEN
' ---------------------------------------------------------------------
' Tue 27-Nov-2018 19:07:29
' Tue 27-Nov-2018 20:24:00
' ---------------------------------------------------------------------


'--------------------------------------------------------
'AFTER THE EFFORT 
'VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
'IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
'HERE VBS DOES SINGLE 
'--------------------------------------------------------


' C:\Users\MATT 04\AppData\Roaming\GoodSync\jobs-groups-options.tic

FILEN_1="\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\GOODSYNC__BACK\_gsdata_\_history_\Profile\jobs-groups-options-2020-0621-204440_2020-0627-194941.tic"


FILEN_2="\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE\GoodSync\Profile\jobs-groups-options.tic-2020-0621-204440-COPY.TIC"


FSO.COPYFILE FILEN_1, FILEN_2

DIM TT_EDITOR
DIM Counter_VAR
DIM Counter_Change_VAR

Set File_1 = fso.OpenTextFile(FILEN_2, ForReading)
TT_EDITOR = File_1.ReadAll


CALL ROUTINE_CREATE_SCRIPT_OF_ALL_JOB

WScript.Quit(0)



SUB ROUTINE_CREATE_SCRIPT_OF_ALL_JOB:

' COUNT THEM
' -------------------------------------------------------------------
X = 1
Do
    T1 = InStr(X, TT_EDITOR, "[n:")
    If T1 > 0 Then
        Counter_VAR = Counter_VAR + 1
    End If

    X = T1 + 2

	IF T1=0 THEN EXIT DO
    T2 = InStr(T1+2, TT_EDITOR, "[n:")
	IF T2=0 AND T1>0 THEN T2=LEN(TT_EDITOR)
    T4 = InStr(T1, TT_EDITOR, "k:1:@|")
	
	IF T4<T2 and T4>0 THEN 
		BLANK_JOBBER=FALSE
		Counter_DUD_JOB_SET = Counter_DUD_JOB_SET + 1
	ELSE
		BLANK_JOBBER=TRUE
	END IF

	T1 = InStr(X-2, TT_EDITOR, "[n:")
	
	T2 = InStr(T1+3, TT_EDITOR, ":")
	N3 = Mid(TT_EDITOR, T1+3, (T2 - T1)-3)
	N4 = Mid(TT_EDITOR, T2+1, CLNG(N3))
	N5 = LPad(CSTR(Counter_VAR),"0",4)+"--"
	
	IF T2=T1 THEN MSGBOX "YY"

	JOB_LINE_FIND_WINDOW_LOW=T1
	JOB_LINE_FIND_WINDOW_HIGH=InStr(T1+2, TT_EDITOR, "[n:")
	IF JOB_LINE_FIND_WINDOW_LOW>0 AND JOB_LINE_FIND_WINDOW_HIGH=0 THEN 
		JOB_LINE_FIND_WINDOW_HIGH=LEN(TT_EDITOR)
	END IF

	JOB_NAME=N4
	
	IF BLANK_JOBBER=TRUE THEN REAL_OR_BLANK="0  "
	IF BLANK_JOBBER=FALSE THEN REAL_OR_BLANK="x--"
	' IMPLY IS FALSE JOB FOR TRUER IS BLANK
	' ---------------------------------------------------------------

	JOB_NAME=N5+REAL_OR_BLANK+JOB_NAME
	
	' JOB_NAME="!AI! && SET Sources[!LDA!]=""" & JOB_NAME & """"

	' --------------------------------
	' SOURCE INFO ABOUT 1ST PATH -- [k:40:/ GET THE K INFO ABOUT TYPE OF PATH 
	' --------------------------------
	TX1 = InStr(T1, TT_EDITOR, "[k:")
	IF TX1>0 THEN 
		TX2 = InStr(TX1+3, TT_EDITOR, ":")
		N3 = Mid(TT_EDITOR, TX1, (TX2 - TX1))
		N5 = TRIM(Replace(N3, "[k:",""))
		N5 = CLNG(N5)
		TD1 = InStr(TX1+1, TT_EDITOR, "[n:") ' CHECK WITHIN FRAME OF ONE JOB NOT OVER TO NEXT
		
		' -----------------------------------------------------------
		' N5 HOLD NUMERIC OF LENGTH STRING PATH
		' GET 1ST PATH
		' -----------------------------------------------------------
		' LINE SOMETIME BEGIN WITH FORWARD SLASH / DISREGARD THAT BIT
		' -----------------------------------------------------------
		START_PATH_POS=TX1+LEN(N5)+4
		PATH_INFO_1=""
		IF N5>0 THEN
			PATH_INFO_1 = Mid(TT_EDITOR,START_PATH_POS,N5)
		END IF
	END IF
	' --------------------------------

	' --------------------------------
	' SOURCE 1ST PATH -- |U:43:/
	' --------------------------------
	TX1 = InStr(T1, TT_EDITOR, "|U:")
	IF TX1>0 THEN 
		TX2 = InStr(TX1+3, TT_EDITOR, ":")
		N3 = Mid(TT_EDITOR, TX1, (TX2 - TX1))
		N5 = TRIM(Replace(N3, "|U:",""))
		N5 = CLNG(N5)
		TD1 = InStr(TX1+1, TT_EDITOR, "[n:") ' CHECK WITHIN FRAME OF ONE JOB NOT OVER TO NEXT
		
		' -----------------------------------------------------------
		' N5 HOLD NUMERIC OF LENGTH STRING PATH
		' GET 1ST PATH
		' -----------------------------------------------------------
		' LINE SOMETIME BEGIN WITH FORWARD SLASH / DISREGARD THAT BIT
		' -----------------------------------------------------------
		START_PATH_POS=TX1+LEN(N5)+4
		PATH_1=""
		IF N5>0 THEN 
			PATH_1 = Mid(TT_EDITOR,START_PATH_POS,N5)
		END IF
		T1=T1+N5+2
	END IF
	' --------------------------------
	' LOCATE AFTER PATH 1
	' --------------------------------
	T4=InStr(T1, TT_EDITOR, "|m:")
	IF T4=0 THEN 
		T4=InStr(T1, TT_EDITOR, "|t")
	END IF
	
	T1=T4

	' --------------------------------
	' IF T1=0 THEN MSGBOX JOB_NAME


	PATH_2=""
	PATH_INFO_2=""
	IF T1>0 THEN 
	' --------------------------------
	' TARGET INFO ABOUT 2ND PATH -- [k:40:/ GET THE K INFO ABOUT TYPE OF PATH 
	' --------------------------------
	TX1 = InStr(T1, TT_EDITOR, "[k:")
	IF TX1>0 THEN 
		TX2 = InStr(TX1+3, TT_EDITOR, ":")
		N3 = Mid(TT_EDITOR, TX1, (TX2 - TX1))
		N5 = TRIM(Replace(N3, "[k:",""))
		N5 = CLNG(N5)
		TD1 = InStr(TX1+1, TT_EDITOR, "[n:") ' CHECK WITHIN FRAME OF ONE JOB NOT OVER TO NEXT
		
		' -----------------------------------------------------------
		' N5 HOLD NUMERIC OF LENGTH STRING PATH
		' GET 1ST PATH
		' -----------------------------------------------------------
		' LINE SOMETIME BEGIN WITH FORWARD SLASH / DISREGARD THAT BIT
		' -----------------------------------------------------------
		START_PATH_POS=TX1+LEN(N5)+4
		PATH_INFO_2=""
		IF N5>0 THEN
			PATH_INFO_2 = Mid(TT_EDITOR,START_PATH_POS,N5)
		END IF
	END IF
	' --------------------------------

	' --------------------------------
	' TARGET 2ND PATH -- |U:43:/
	' --------------------------------
	TX1 = InStr(T1, TT_EDITOR, "|U:")
	IF TX1>0 THEN 
		TX2 = InStr(TX1+3, TT_EDITOR, ":")
		N3 = Mid(TT_EDITOR, TX1, (TX2 - TX1))
		N5 = TRIM(Replace(N3, "|U:",""))
		N5 = CLNG(N5)
		TD1 = InStr(TX1+1, TT_EDITOR, "[n:") ' CHECK WITHIN FRAME OF ONE JOB NOT OVER TO NEXT
		
		' -----------------------------------------------------------
		' N5 HOLD NUMERIC OF LENGTH STRING PATH
		' GET 1ST PATH
		' -----------------------------------------------------------
		' LINE SOMETIME BEGIN WITH FORWARD SLASH / DISREGARD THAT BIT
		' -----------------------------------------------------------
		START_PATH_POS=TX1+LEN(N5)+4
		PATH_2=""
		IF N5>0 THEN 
			PATH_2 = Mid(TT_EDITOR, START_PATH_POS,N5)
		END IF
	END IF
	END IF
	' --------------------------------

	STRING_SET=JOB_NAME + VBCRLF +PATH_INFO_1 + VBCRLF + PATH_1 + VBCRLF + PATH_INFO_2 + VBCRLF +PATH_2 + VBCRLF + "----------------------------------------" + VBCRLF
	' MSGBOX STRING_SET
	I_X=I_X+STRING_SET
	Counter_VAR = Counter_VAR + 1

	' --------------------------------
	' LOCATE NEXT JOB AR
	' --------------------------------
	IF T1>0 THEN 
	T4=InStr(T1, TT_EDITOR, "|]|")         ' END EACH JOB -- END OF ALL MAIN JOB HERE -- |]|}g()d(
	IF T4=0 THEN 
		T4=InStr(T1, TT_EDITOR, "|]"+VBLF) ' END TYPE OF INDEX SYSTEM -- WRITE END JOB WAY
	END IF
	IF T4>0 THEN 
		T1=T4-10
	ELSE
		T1=0
	END IF
	END IF

	X = T1 + 2
		
Loop Until T1 = 0



' FILEN_2="C:\SCRIPTER\NOTEPAD TALK\TEXT 2020-06-26 __ GS-STRING.TXT"
' Set objFile = FSO.CreateTextFile(FILEN_2,True)
' objFile.Close
' Set File_2 = fso.OpenTextFile(FILEN_2, ForWriting)
' File_2.Write I_X



FILEN_1="C:\GOODSYNC JOB CONTROL\01 VB OUTPUT SCRIPT CONTROL\SCRIPTOR OF 7G_D_GS2GO__2020-0621-204440.TXT"
FILEN_2="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 08-GOODSYNC EDITOR_4 MAKE SCRIPT OF JOB_7G_D_GS2GO_2020-0621-204440.TXT"

FOR R=1 TO 2
	IF R=1 THEN FILE_NAME=FILEN_1
	IF R=2 THEN FILE_NAME=FILEN_2
	Set objFile = FSO.CreateTextFile(FILE_NAME,True)
	objFile.Close
	Set File_2 = fso.OpenTextFile(FILE_NAME, ForWriting)
	File_2.Write I_X
NEXT 




MsgBox "Done Total " & Counter_VAR & " Proecesses"+VBCRLF+"Total Changes "& Counter_Change_VAR,vbMsgBoxSetForeground,Wscript.ScriptName

WScript.Quit(0)


END SUB


'--------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function