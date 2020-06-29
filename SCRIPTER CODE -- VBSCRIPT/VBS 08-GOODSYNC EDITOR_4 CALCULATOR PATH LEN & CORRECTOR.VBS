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
' SPAWN OF 
' -------------------------------------------------------------------
' VBS 08-GOODSYNC EDITOR_4 MAKE SCRIPT OF JOB.VBS
' -------------------------------------------------------------------
' SPAWN OF 
' -------------------------------------------------------------------
' VBS 08-GOODSYNC EDITOR_4 EXTRACT WORK JOB & PATH.VBS
' -------------------------------------------------------------------
' WORK BOSH
' Sun 28-Jun-2020 04:18:40
' Sun 28-Jun-2020 06:49:04 -- 2 HOUR 30 MINUTE
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' HERE
' -------------------------------------------------------------------
' VBS 08-GOODSYNC EDITOR_4 CALCULATOR PATH LEN & CORRECTOR.VBS
' DESCRIPTION 
' -------------------------------------------------------------------
' FOR EVENT WHEN SEARCH REPLACE RENAME
' RE-DO CALCULATOR THE INDEX STRING LENGTH
' -------------------------------------------------------------------
' WORK BOSH
' Sun 28-Jun-2020 07:00:00
' Sun 28-Jun-2020 07:07:40 -- 
' -------------------------------------------------------------------






' -------------------------------------------------------------------
DIM CUR_USER
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

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

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

CUR_USER = CreateObject("WScript.Network").UserName

' C:\Users\MATT 04\AppData\Roaming\GoodSync\jobs-groups-options.tic

' FILEN_1="\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\GOODSYNC__BACK\_gsdata_\_history_\Profile\jobs-groups-options-2020-0621-204440_2020-0627-194941.tic"

' FILEN_1="\\4-ASUS-GL522VW\4_ASUS_GL522VW_02_D_DRIVE\GoodSync\Profile\jobs-groups-options.tic"

' FILEN_2="\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\GOODSYNC__BACK\Profile\jobs-groups-options.tic--2020-0628-COPY.TIC"


' FSO.COPYFILE FILEN_1, FILEN_2

DIM TT_EDITOR
DIM Counter_VAR
DIM Counter_Change_VAR


CALL ROUTINE_CREATE_SCRIPT_OF_ALL_JOB

WScript.Quit(0)



SUB ROUTINE_CREATE_SCRIPT_OF_ALL_JOB:


DIM FILEN(4)
DIM FILEN_BACKER()
REDIM FILEN_BACKER(UBOUND(FILEN))



FILEN(1)="\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE\GoodSync\Profile\jobs-groups-options.tic"
FILEN(2)="\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\GOODSYNC__BACK\_gsdata_\_history_\Profile\jobs-groups-options-2020-0621-204440_2020-0627-194941.tic"
FILEN(3)="\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE\Users\MATT 04\AppData\Local\GoodSync\jobs-groups-options.tic"
FILEN(4)="\\4-ASUS-GL522VW\4_ASUS_GL522VW_02_D_DRIVE\GoodSync\Profile\jobs-groups-options.tic"

' -------------------------------------------------------------------
' HEAVY OF BACKER __ 
' EVERY FILE WRITE AS TIME DATE FILENAME OF EACH OWN DIRECTORY
' -------------------------------------------------------------------
FOR R=1 TO UBOUND(FILEN)
	FILEN_BACKER(R) = FILEN(R) + "__" + GET_TIME_DATE_AND_FORMAT + ".tic"
NEXT
' -------------------------------------------------------------------
FOR R=1 TO UBOUND(FILEN_BACKER)
	FSO.COPYFILE FILEN(R), FILEN_BACKER(R)
NEXT 
' -------------------------------------------------------------------

Counter_Change_VAR=0


FOR R=1 TO UBOUND(FILEN)

Set File_1 = fso.OpenTextFile(FILEN(R), ForReading)
TT_EDITOR = File_1.ReadAll
TT_EDITOR_2=TT_EDITOR
' COUNT THEM
' -------------------------------------------------------------------
Counter_VAR=0
T1=1
I4=0
Do
	IF T1<1 THEN EXIT DO

    T2 = InStr(T1, TT_EDITOR, "n:")

	T3 = InStr(T2+2, TT_EDITOR, ":")
	N3 = Mid(TT_EDITOR, T2+2, (T3 - T2)-2)
	N4 = Mid(TT_EDITOR, T3+1, CLNG(N3))

	N5= InStr(T3, TT_EDITOR, "|")
	N8=	Mid(TT_EDITOR, T3+1, (N5 - T3 - 1))

	' ---------------------------------------------------------------
	' N3 CONTAIN LEN STRING THAT GO TO BE N4 -- JOB NAME
	' N4 -- JOB NAME STRING PROVIDED BY N3 LENGTH OF AH 
	' ---------------------------------------------------------------

	' ---------------------------------------------------------------
	' IN CASE N3 WRONG BY RENAME SEARCH REPLACE
	' ---------------------------------------------------------------
	' N5 OF END MARK WHERE STRING ARE
	' N5 POSITION WHEE STRING END
	' N8 STING REGARDLESS IF LEN FROM N3 AND FIND END MARKER INSTEAD
	' ---------------------------------------------------------------

	' SEARCH REPLACE INFO WITH CORRECTOR
	' REQUIRE BUILD INITIAL STRING -- ADD PATH -- OF WHAT GATHER BY END MARKER
	' WE GET N8 FOR THAT 
	' N8 IS CORRECT PATH WITH LEN ON
	' BUT WE WANT PATH AS LEN ISSUE WRONG
	' MUST BE DIFFERENT
	' REALLY IF RENAME BEEN DONE 
	' NONE OPTION 
	' IT ONLY NUMERIC LENGTH VARIABLE -- NOT RELOCATE ANY OF THAT PROFILE TT_EDITOR
	' USE PATH NAME TWO

	IF N4<>N8 THEN 
		N10=CSTR(LEN(N8))
		TT_EDITOR = Replace(TT_EDITOR, "n:"+N3+":"+N8,"n:"+N10+":"+N8)
		I=""
		I= I + CSTR(R) & " OF " & CSTR(UBOUND(FILEN)) & VBCR & VBCR
		I= I + FILEN(R) & VBCR  & VBCR

		I= I + N3 & VBCR & N4 & VBCR & LEN(N8) & VBCR & N8  & VBCR
		I= I + VBCR & "n:"+N3+":"+N8
		I= I + VBCR & "n:"+N10+":"+N8
		Counter_VAR=Counter_VAR+1
		Counter_Change_VAR=Counter_Change_VAR+1
		IF I4<2 THEN 
			MSGBOX I
			I4=I4+1
		END IF
	END IF
	T1=InStr(T2+2, TT_EDITOR, "n:")-4
	IF T1=0 THEN EXIT DO
LOOP

MsgBox "Done " & R & " OF " & UBOUND(FILEN) & VBCR & VBCR & FILEN(R) & VBCR & VBCR & "Proecessor -- " & Counter_VAR & VBCR & "Total Changes -- " &  Counter_Change_VAR , vbMsgBoxSetForeground , Wscript.ScriptName

FOR R2=4 TO 4
	IF TT_EDITOR_2<>TT_EDITOR THEN 
		FILE_NAME=FILEN(R) 
		Set objFile = FSO.CreateTextFile(FILE_NAME,True)
		objFile.Close
		Set File_2 = fso.OpenTextFile(FILE_NAME, ForWriting)
		File_2.Write TT_EDITOR
		MSGBOX "WRITE DONE" & VBCR & FILE_NAME
	END IF
NEXT 


NEXT

WScript.Quit(0)

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


' MsgBox "Done Total " & Counter_VAR & " Proecesses"+VBCRLF+"Total Changes "& Counter_Change_VAR,vbMsgBoxSetForeground,Wscript.ScriptName



WScript.Quit(0)


END SUB


'--------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function


Function GET_TIME_DATE_AND_FORMAT()
	VAR_NOW=NOW
	YY=Year(VAR_NOW)
	MM=MONTH(VAR_NOW)
	MMM = MonthName(MONTH(VAR_NOW),TRUE)
	DD=Day(VAR_NOW)
	HH24=Hour(VAR_NOW)
	MM=Minute(VAR_NOW)
	SS=Second(VAR_NOW)
	HH12=Hour(VAR_NOW)
	IF HH12>12 THEN 
		HH12=HH12-12
		AM_PM="PM"
	ELSE
		AM_PM="AM"
	END IF 
	IF HH12=12 THEN 
		AM_PM="PM"
	END IF 

	DDD=WeekdayName(Weekday(VAR_NOW),TRUE)

	' -------------------------------------------
	' 2020-Jun-29 -- 12:07:25 - 12:07:25-PM - Mon
	' -------------------------------------------
	i_TIME=""
	i_TIME=i_TIME+LPad(YY, "0", 4)+"-"
	i_TIME=i_TIME+MMM+"-"
	i_TIME=i_TIME+LPad(DD, "0", 2)
	i_TIME=i_TIME+" -- "
	i_TIME=i_TIME+LPad(HH24, "0", 2)+":"
	i_TIME=i_TIME+LPad(MM, "0", 2)+":"
	i_TIME=i_TIME+LPad(SS, "0", 2)
	i_TIME=i_TIME+" - "
	i_TIME=i_TIME+LPad(HH12, "0", 2)+":"
	i_TIME=i_TIME+LPad(MM, "0", 2)+":"
	i_TIME=i_TIME+LPad(SS, "0", 2)+"-"+AM_PM
	i_TIME=i_TIME+" - "
	i_TIME=i_TIME+DDD
	' -------------------------------------------
	' 2020-Jun-29 -- 12:07:25 - 12:07:25-PM - Mon
	' -------------------------------------------


	' -------------------------------------------
	' 2020-12-29 12-12-19
	' -------------------------------------------
	i_TIME=""
	i_TIME=i_TIME+LPad(YY, "0", 4)+"-"
	i_TIME=i_TIME+LPad(MM, "0", 2)+"-"
	i_TIME=i_TIME+LPad(DD, "0", 2)
	i_TIME=i_TIME+" "
	i_TIME=i_TIME+LPad(HH24, "0", 2)+"-"
	i_TIME=i_TIME+LPad(MM, "0", 2)+"-"
	i_TIME=i_TIME+LPad(SS, "0", 2)
	' -------------------------------------------
	' 2020-12-29 12-12-19
	' -------------------------------------------

	GET_TIME_DATE_AND_FORMAT=i_TIME
End Function


