

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

SET objArgs = Wscript.Arguments
FOR EACH strArg In objArgs

	' PARAM = PARAM + """"+REPLACE(strArg,"*"," ")+"""" +" "
	PARAM = PARAM + """"+strArg+"""" +" "
NEXT

IF PARAM="" THEN 
	Wscript.QUIT
END IF


' "--JOBNAME--" "VB VB-NT 01 EXE 4G 1X _001_ EXE" "--RESULT--" "" "--ERROR--" "1" "--CONFLICT--" "0" "--LOGPATH--" "D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@4-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT\2020-0726-082045-4-ASUS-GL522VW-VB VB-NT 01 EXE 4G 1X _001_ EXE.log" 
' "--JOBNAME--" "VB VB-NT 01 EXE 7G 4G _001_ EXE" "--RESULT--" "" "--ERROR--" "1" "--CONFLICT--" "0" "--LOGPATH--" "D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@7-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT-00_BEST_VB_01\2020-0726-094314-4-ASUS-GL522VW-VB VB-NT 01 EXE 7G 4G _001_ EXE.log" 



XX=INSTR(PARAM,"--LOGPATH--")
IF XX>0 THEN 
	LOGPATH=MID(PARAM,XX+LEN("--LOGPATH--")+2)
	LOGPATH=REPLACE(LOGPATH,"""","")
END IF

IF FSO.FILEEXISTS(LOGPATH) = TRUE THEN 
	Set FILE_1 = FSO.OpenTextFile(LOGPATH, ForReading)
	TT_EDITOR = FILE_1.ReadAll
	' MSGBOX TT_EDITOR
END IF

IF FSO.FILEEXISTS(LOGPATH) = TRUE THEN 
	' ---------------------------------------------
	Set WSHShell = CreateObject("WScript.Shell")
	' WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+LOGPATH+""""
	Set WSHShell = Nothing
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


Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

IF COMPUTERNAME_1<>UCASE(ComputerName) THEN 
	COMPUTERNAME_REMOTE=COMPUTERNAME_1
END IF
IF COMPUTERNAME_2<>UCASE(ComputerName) THEN 
	COMPUTERNAME_REMOTE=COMPUTERNAME_1
END IF

XXVV=XXVV+VBCR+COMPUTERNAME_REMOTE

' SO FAR SO GOOD
' SEND THE INFO TO REMOTE COMPUTER
' ABOUT KILL PROCESS SOMETHING

COMPUTERNAME_REMOTE_UNBDERLINE=REPLACE(COMPUTERNAME_REMOTE,"-","_")

PATH_1=""
PATH_1=PATH_1+"\\"+COMPUTERNAME_REMOTE+"\"+COMPUTERNAME_REMOTE_UNBDERLINE+"_01_C_DRIVE"
PATH_1=PATH_1+"\SCRIPTOR DATA\SCRIPTER CODE -- VBSCRIPT\VBS 79-GOODSYNC SCRIPT VB EXE_VBS_DATA_#NFS_EX.TXT"

' THAT SHOULD DO
' WRITE PROCESS EXE NAME

IF FSO.FILEEXISTS(PATH_1) = TRUE THEN 
	FSO.DELETEFILE(PATH_1)
END IF

SET OBJFILE = FSO.CreateTextFile(PATH_1,True)
OBJFILE.Close

WRITE_FILE_INFO=JOBNAME + VBCRLF + EXE_NAME + VBCRLF

SET FILE_2 = FSO.OpenTextFile(PATH_1, ForWriting)
FILE_2.Write WRITE_FILE_INFO


Wscript.QUIT


