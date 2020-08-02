' ---------------------------------------------------------------
' NEW FABBY CODE TO ASSIST GOODSYNC 
' WHERE MY VISUAL BASIC CODE VB6 __ THE APP ARE KILL TO ALLOW GOODSYNC COPY OVER
' AT REMOTE LOCATION
' THE SCRIPT RERUN AGAIN AFTER TOO
' A BIT LATER THAT LAST BIT
' DELAY REMOTE SENSOR REQUIRE
' WRITE TIME
' Sun 26-Jul-2020 08:06:11
' Sun 26-Jul-2020 10:48:00 -- 2 HOUR 41 MINUTE
' Sun 26-Jul-2020 14:43:42 -- 6 HOUR 37 MINUTE
' ---------------------------------------------------------------
' HERE WORK WITH HER
' Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
' ROUTINE ---- TIMER_GOODSYNC_DELETE_PROCESS_VIA_VBS:
' ---------------------------------------------------------------
' AND HERE
' C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE.VBS
' ---------------------------------------------------------------

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Const ForReading = 1, ForWriting = 2, ForAppending = 8

DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")

' -------------------------------------------------------------------
DIM WSHShell
SET WSHShell = CreateObject("WScript.Shell")
SET objShell = Wscript.CreateObject("WScript.Shell")
CUR_USER = CreateObject("WScript.Network").UserName
COMPUTERNAME = objShell.ExpandEnvironmentStrings("%computername%")
' -------------------------------------------------------------------
' 01 OF 02 
' SOME CONFLICT DISAGREE-ER IN CODE STYLE 
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
' -------------------------------------------------------------------
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
		
		IF iMyPID<>PID_Script THEN 
			' MSGBOX iMyPID & VBCR & PID_Script
			objProcess.Terminate() 
			objShell.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
		END IF
		
		'IF i2 > 0 THEN ' MY PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			' IF lngMyProcessId<>iMyPID THEN 
				' I= "ERROR MISMATCH TRUE MYPID RESULT "
				' MSGBOX I +VBV+ CSTR(iMyPID) +VBV+ CSTR(lngMyProcessId)
			' -------------------------------------------------------
			' IF iMyPID<>PID_Script THEN 
				' MSGBOX iMyPID & VBCR & PID_Script
				' ' MSGBOX PID_Script
				' objProcess.Terminate() 
				' objShell.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
				' ' MSGBOX CSTR(iMyPID) +VBV+ strScriptName +VBV+ CSTR(PID_Script)  +VBV+ Wscript.ScriptName
			' END IF
		' END IF
    END IF
NEXT
' -------------------------------------------------------------------



' -------------------------------------------------------------------
' GOODSYNC_GSYNC="D:\GoodSync\x64\gsync.exe"
' JOBNAME="VB VB-NT 01 EXE 7G 4G _001_ EXE"
' Set WSHShell = CreateObject("WScript.Shell")
' WSHShell.Run """" + GOODSYNC_GSYNC + """" + " sync " + """"+JOBNAME+""""
' Set WSHShell = Nothing
' WSCRIPT.QUIT
' -------------------------------------------------------------------


SET objArgs = Wscript.Arguments
FOR EACH strArg In objArgs

	' PARAM = PARAM + """"+REPLACE(strArg,"*"," ")+"""" +" "
	PARAM = PARAM + """"+strArg+"""" +" "
NEXT

IF PARAM="" THEN 
	Wscript.QUIT
END IF


' -------------------------------------------------------------------
' "--JOBNAME--" "VB VB-NT 01 EXE 4G 1X _001_ EXE" "--RESULT--" "" "--ERROR--" "1" "--CONFLICT--" "0" "--LOGPATH--" "D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@4-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT\2020-0726-082045-4-ASUS-GL522VW-VB VB-NT 01 EXE 4G 1X _001_ EXE.log" 
' -------------------------------------------------------------------
' "--JOBNAME--" "VB VB-NT 01 EXE 7G 4G _001_ EXE" "--RESULT--" "" "--ERROR--" "1" "--CONFLICT--" "0" "--LOGPATH--" "D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@7-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT-00_BEST_VB_01\2020-0726-094314-4-ASUS-GL522VW-VB VB-NT 01 EXE 7G 4G _001_ EXE.log" 
' -------------------------------------------------------------------



' -------------------------------------------------------------------
' LINE TO ENTER AT GOODSYNC SCRIPTOR -- OF POST SYNC -- EXAMPLE
' -------------------------------------------------------------------
' "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE.VBS" --JOBNAME-- %JOBNAME% --RESULT-- %RESULT% --ERROR-- %ERRORS% --CONFLICT-- %CONFLICTS% --LOGPATH-- %LOGPATH%
' -------------------------------------------------------------------
' Wed 29-Jul-2020 04:54:24
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' MORE WORK DO
' -------------------------------------------------------------------
' 001
' IDENTIFY PROCCESS.EXE NAME OPERATION BASE ON
' -------------------------------------------------------------------
' 002
' CLOSE CERTAIN JOB FRIENDLY WAY -
' DOUBLE CHECK WAITER
' -------------------------------------------------------------------
' 003
' STOP REPEAT FAIL KILL
' -------------------------------------------------------------------





XX=INSTR(PARAM,"--LOGPATH--")
IF XX>0 THEN 
	LOGPATH=MID(PARAM,XX+LEN("--LOGPATH--")+2)
	LOGPATH=REPLACE(LOGPATH,"""","")
END IF

' THE SCRIPT LOG NOT WRITE FROM GOODSYNC AND USER 
' THE NOWAIT FOR LOAD SCRIPT OPTION
' -------------------------------------------------------------------
IF FSO.FILEEXISTS(LOGPATH) = TRUE THEN 
	Set FILE_1 = FSO.OpenTextFile(LOGPATH, ForReading)
	TT_EDITOR = FILE_1.ReadAll
	FILE_1.CLOSE
	' MSGBOX TT_EDITOR
END IF

IF FSO.FILEEXISTS(LOGPATH) = TRUE THEN 
	' ---------------------------------------------
	'Set WSHShell = CreateObject("WScript.Shell")
	' WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+LOGPATH+""""
	'Set WSHShell = Nothing
	' ---------------------------------------------
	' LOSE ACCESS TO DISPLAY NOTEPAD++ AND EXPLORER
	' NETWORK -------------------------------------
	' AND SUDDEN COME BACK
	' AFTER NEXT LINE CREATOR
	' NETWORK -------------------------------------
	' DELAY GOODSYNC BEFORE WRITE LOG POSSIBLE
	' ---------------------------------------------
	' DO READ BEFORE
	' ---------------------------------------------
	' AND THEN SPORADIC -- READ INFO VBS OKAY
	' REM OUT NOTEPAD -- ALL DO
	' ---------------------------------------------
END IF

XX=INSTR(PARAM,"--JOBNAME--")
IF XX>0 THEN
	JOBNAME=MID(PARAM,XX+LEN("--JOBNAME--")+3)
	JOBNAME=MID(JOBNAME,1,INSTR(JOBNAME,""""))
	JOBNAME=REPLACE(JOBNAME,"""","")
END IF


DIM NETNAME(7)
i = i + 1: NETNAME(i) = "1-ASUS-X5DIJ"
i = i + 1: NETNAME(i) = "2-ASUS-EEE"
i = i + 1: NETNAME(i) = "3-LINDA-PC"
i = i + 1: NETNAME(i) = "4-ASUS-GL522VW"
i = i + 1: NETNAME(i) = "5-ASUS-P2520LA"
i = i + 1: NETNAME(i) = "7-ASUS-GL522VW"
i = i + 1: NETNAME(i) = "8-MSI-GP62M-7RD"

' WAIT FOR LOG FILE ARRIVE SEEM DELAY CREATE FROM REQUEST SCRIPT FILE RUN

' DEBUGGER KITT USER WHEN FIND NOWAIT OPTION
COUNTER_LOOP=0
DO 
	IF FSO.FILEEXISTS(LOGPATH) = TRUE THEN   EXIT DO
	WScript.Sleep 100
	COUNTER_LOOP=COUNTER_LOOP+1
	IF COUNTER_LOOP>100 THEN                 EXIT DO
LOOP

IF FSO.FILEEXISTS(LOGPATH) = FALSE THEN 
	MSGBOX "TT FSO.FILEEXISTS(LOGPATH) = FALSE"+VBCR+VBCR+LOGPATH
	WSCRIPT.QUIT
END IF

SET FILE_1 = FSO.OpenTextFile(LOGPATH)
Do Until FILE_1.AtEndOfStream
	READ_LINE=FILE_1.ReadLine
	
	IF MM=0 THEN 
		I="Cannot delete file: Access is denied."
		MM=INSTR(READ_LINE,I)
	END IF
	IF MM=0 THEN 
		I="The process cannot access the file because it is being used by another process."
		MM=INSTR(READ_LINE,I)
	END IF
	I="Copy Over '"
	M1=INSTR(READ_LINE,I)
	IF M1>0 THEN 
		FOR R=1 TO 7
			IF INSTR(UCASE(READ_LINE),UCASE(NETNAME(R)))>0 THEN 
				COMPUTERNAME_1=UCASE(NETNAME(R))
			END IF
		NEXT 
		FOR R=1 TO 7
			IF INSTR(UCASE(READ_LINE),UCASE(NETNAME(R)))>0 AND NETNAME(R)<>COMPUTERNAME_1 THEN 
				COMPUTERNAME_2=UCASE(NETNAME(R))
			END IF
		NEXT 
		
		I=UCASE(".EXE'")
		M2=INSTR(UCASE(READ_LINE),I)
		IF M2>0 THEN 
			M4=INSTRREV(UCASE(READ_LINE),"/")
			EXE_NAME=MID(READ_LINE,M4+1)
			M2=INSTR(UCASE(EXE_NAME),I)
			EXE_NAME=MID(EXE_NAME,1,M2+3)
		END IF
	END IF

	IF MM>0 THEN
		XXVV=READ_LINE+VBCR+COMPUTERNAME_1+VBCR+COMPUTERNAME_2+VBCR+EXE_NAME
	END IF
	MM=0
Loop
FILE_1.CLOSE

Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

IF COMPUTERNAME_1<>UCASE(ComputerName) THEN 
	COMPUTERNAME_REMOTE=COMPUTERNAME_1
END IF
IF COMPUTERNAME_2<>UCASE(ComputerName) THEN 
	COMPUTERNAME_REMOTE=COMPUTERNAME_2
END IF

XXVV=XXVV+VBCR+COMPUTERNAME_REMOTE

' -------------------------------------------------------------------
' SO FAR SO GOOD
' SEND THE INFO TO REMOTE COMPUTER
' ABOUT KILL PROCESS SOMETHING
' -------------------------------------------------------------------

COMPUTERNAME_REMOTE_UNBDERLINE=REPLACE(COMPUTERNAME_REMOTE,"-","_")

IF COMPUTERNAME_REMOTE="" THEN WSCRIPT.QUIT

PATH_1=""
PATH_1=PATH_1+"\\"+COMPUTERNAME_REMOTE+"\"+COMPUTERNAME_REMOTE_UNBDERLINE+"_01_C_DRIVE"
PATH_1=PATH_1+"\SCRIPTOR DATA\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE_VBS_DATA_#NFS_EX.TXT"

' -------------------------------------------------------------------
' THAT SHOULD DO
' WRITE PROCESS EXE NAME
' -------------------------------------------------------------------

IF FSO.FILEEXISTS(PATH_1) = TRUE THEN 
	FSO.DELETEFILE(PATH_1)
END IF

FULLPATH=MID(PATH_1,1,INSTRREV(PATH_1,"\"))
IF FULLPATH<>"" THEN 
	BuildFullPath FULLPATH
END IF

DIM OBJFILE

ON ERROR RESUME NEXT
ERR.CLEAR 
IF FSO.FILEEXISTS(PATH_1) = FALSE THEN 
	ERR.CLEAR 
	SET OBJFILE = FSO.CreateTextFile(PATH_1)
	' ---------------------------------------------------------------
	' GET ERROR -- SET OBJFILE = FSO.CreateTextFile(PATH_1)
	' Permission denied
	' FOR AGER WHEN BuildFullPath __ CREATE FOLDER WHERE FILE TO BE FOLDER AS NAME FILE
	IF ERR.NUMBER >0 THEN MSGBOX "LINE 283" & VBCR & "ERROR -- SET OBJFILE = FSO.CreateTextFile(PATH_1)" & VBCR & ERR.DESCRIPTION & VBCR & VBCR & PATH_1
	' ---------------------------------------------------------------
	ERR.CLEAR
	OBJFILE.Close
	IF ERR.NUMBER >0 THEN MSGBOX "LINE 286" & VBCR & "ERROR -- OBJFILE.Close"
END IF
IF ERR.NUMBER>0 THEN 
	MSGBOX "LINE 276" & VBCR & ERR.DESCRIPTION & VBCR & VBCR & PATH_1
END IF
ON ERROR GOTO 0

' -------------------------------------------------------------------
' Object required
' \\\_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE_VBS_DATA_#NFS_EX.TXT
' Object required
' \\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE_VBS_DATA_#NFS_EX.TXT
' -------------------------------------------------------------------


WRITE_FILE_INFO=JOBNAME + VBCRLF + EXE_NAME + VBCRLF

ON ERROR RESUME NEXT
SET FILE_2 = FSO.OpenTextFile(PATH_1, ForWriting)
FILE_2.Write WRITE_FILE_INFO
ERR.CLEAR
FILE_2.Close
IF ERR.NUMBER >0 THEN MSGBOX "LINE 306" & VBCR & "ERROR -- FILE_2.Close"
IF ERR.NUMBER >0 THEN 
	MSGBOX "LINE 308" & VBCR & ERR.DESCRIPTION & VBCR & VBCR & PATH_1
END IF
ON ERROR GOTO 0
' -------------------------------------------------------------------
' Script:	C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE.VBS
' Line:	263
' Char:	1
' Error:	Permission denied
' Code:	800A0046
' Source: 	Microsoft VBScript runtime error
' -------------------------------------------------------------------

IF INSTR(LOGPATH,"D:\GOODSYNC\Profile\")>0 THEN 
	GOODSYNC_GSYNC="D:\GoodSync\x64\gsync.exe"
END IF

' -------------------------------------------------------------------
' @SET GoodSync_GSync="C:\Program Files\Siber Systems\GoodSync\gsync.exe"
' SET Sources[!LDA!]="0001 C GOODSYNC 2 GO -- D TO D __ 7G 4G _ JOB SET"
' %GoodSync_GSync% sync !Sources_02[%%a]!>!FILENAME_ER!
' -------------------------------------------------------------------

COUNTER_LOOP=0
DO 
	IF FSO.FILEEXISTS(PATH_1) = FALSE THEN   EXIT DO

	WScript.Sleep 1000
	COUNTER_LOOP=COUNTER_LOOP+1
	IF COUNTER_LOOP>100 THEN                 EXIT DO
LOOP


IF FSO.FILEEXISTS(PATH_1) = FALSE THEN

	' ---------------------------------------------------------------
	' Script:	C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE.VBS
	' Line:	274  ---- HERE ---- WSHShell.Run """" + GOODSYNC_GSYNC + """" + " sync " + """"+JOBNAME+""""
	' Char:	2
	' Error:	The system cannot find the file specified. 
	' Code:	80070002
	' Source: 	(null)
	' ---------------------------------------------------------------
	' SET ERROR TRAP
	' AFTER JOB EXIT COMPLETE FINE THIS STATE
	' ---------------------------------------------------------------
	
	ON ERROR RESUME NEXT
	' MSGBOX JOBNAME
	Set WSHShell = CreateObject("WScript.Shell")
	WSHShell.Run """" + GOODSYNC_GSYNC + """" + " sync " + """"+JOBNAME+""""
	Set WSHShell = Nothing
	ON ERROR GOTO 0

END IF


' -------------------------------------------------------------------
' -------------------------------------------------------------------
Wscript.QUIT
' -------------------------------------------------------------------
' -------------------------------------------------------------------

Sub BuildFullPath(ByVal FullPath)
	
	IF FullPath="" THEN EXIT SUB
	
	IF NOT FSO.FileExists(FullPath) Then
	IF NOT FSO.FolderExists(FullPath) Then
		ON ERROR RESUME NEXT
		BuildFullPath FSO.GetParentFolderName(FullPath)
		ERR.CLEAR
		FSO.CreateFolder FullPath
		IF ERR.NUMBER>0 THEN MSGBOX "LINE 373" & VBCR & "ERROR -- FSO.CreateFolder FullPath" & VBCR & ERR.DESCRIPTION & VBCR & VBCR & FullPath
		ON ERROR GOTO 0
	End If
	End If
End Sub

