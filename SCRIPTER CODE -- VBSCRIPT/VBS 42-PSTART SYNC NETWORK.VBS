'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 42-PSTART SYNC NETWORK.VBS
'# __ 
'# __ VBS 42-PSTART SYNC NETWORK.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Fri 22-Mar-2019 10:31:08 ]
'# END FINAL TIME [ Fri 22-Mar-2019 11:50:00 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' NEW CODE THE PStart.xml
' HARD TO SYNC AS GOT A COUNTER FOR EVERY USE
' CONFLICT TOO HARD OVER NETWORK
' SO SYNC WITH OWN METHOD
' MOST HIGHER COUNT VERSION IS SYNC-ER 
' IN MY CODE
'
' ALL DONE AND DUSTED 
'
' NEXT PROBLEM WHILE THE CODE IS RUNNER PSTART.EXE
' THE TIME COUNTER RUN TIME KEEP UPDATE 
' SO NOT TO LEAVE RUNNER IN ICON TASK-BAR ICON TRAY BAR
'
' AND THE ADDITIONAL TASK OF ADD CODING FOR AUTOHOTKEYS
' TO KILL PROCESS PSTART.EXE IF NOT BEEN USED FOR A WHILE LIKE 20 MIN
' TOOK CODER FRIDAY MORNING UP TO Fri 22-Mar-2019 12:58:00
'
' ALSO DURING WAS A MENU OPTION IN ANOTHER CODE
' TO LET ME LOOK AT CLIPBOARD IN HEX VIEWER FOR VBCRLF LINE ENDING
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 22-Mar-2019 10:31:08 ]
'# END FINAL TIME [ Fri 22-Mar-2019 11:50:00 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF ARE
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
		strScriptName = Left(strScriptName, Len(strScriptName) - 1)                             
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		else
			i1=i1+1
		END IF
		IF i1 > 0 THEN 
				' Wscript.Quit
		END IF  
		IF i2 > 1 THEN 
				Wscript.Quit
		END IF  
    END IF  
NEXT

Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'--------------------------------------------------------
' AFTER THE EFFORT 
' VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
' IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
' HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

Dim Array_FileName()
Dim Array_NETPATH_01()
Dim Array_NETPATH_02()
Dim Array_PSTART_COUNTER()
ReDim Preserve Array_FileName(50)
ReDim Preserve Array_NETPATH_01(50)
ReDim Preserve Array_NETPATH_02(50)
ReDim Preserve Array_PSTART_COUNTER(50)

CALL READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY

FOR RI=1 TO UBOUND(Array_FileName)
	Array_FileName(RI)=Array_FileName(RI)+"_01_C_DRIVE\PStart\PStart.xml"
NEXT

CALL READ_PSTART_FOR_BIGGER

' SELECT BIGGER COUNTER FROM ALL NETWORK PStart.xml
FOR RI=1 TO UBOUND(Array_PSTART_COUNTER)
	T_VAR=Array_PSTART_COUNTER(RI)
	IF T_VAR<>"" THEN 
		R_COUNT=MID(T_VAR,INSTR(T_VAR,"<count>")+7)
		T_VAR=R_COUNT
		R_COUNT=MID(T_VAR,1,INSTR(T_VAR,"</count>")-1)
		IF R_COUNT>R_COUNT_REMEMBER THEN 
			R_COUNT_REMEMBER=R_COUNT
			R_COUNT_REMEMBER_ID=RI
		END IF
	END IF
	' <count>4834</count>
NEXT

' -------------------------------------------------------------------
' PUT BIGGER ON THEM ALL NETWORK PStart.xml BUT NOT OWN ONLY IF LESS COUNTER 
' I SEE A PROBLEM ANY 1 NUDGE UP COUNTER MIGHT BE CONFLICT AGAIN
' TO GET OVER PROBLEM ANY STRING NOT MATCH AND PRIORITY TO 1ST HIGHER COUNTER
' BUT STILL A PROBLEM NEVER A SYNC CONFLICT AGAIN
' WHICH IS NOT TRUE FOR CONFLICT 
' BEST WE GET
' NOT TRUE LEAVE THE CONFLICT ONE AS 1 NUDGE UP ALONE
' AND WAIT IF ANY HIGHER 
' MAKE DO ADD TRUE CONFLICT AGAIN
' -------------------------------------------------------------------
FOR RI=1 TO UBOUND(Array_FileName)
	T_VAR=Array_PSTART_COUNTER(RI)
	IF T_VAR<>"" THEN 
		R_COUNT=MID(T_VAR,INSTR(T_VAR,"<count>")+7)
		T_VAR=R_COUNT
		R_COUNT=MID(T_VAR,1,INSTR(T_VAR,"</count>")-1)
		IF R_COUNT<R_COUNT_REMEMBER THEN 
			IF Array_FileName(R_COUNT_REMEMBER_ID)<> Array_FileName(RI) THEN 
				IF FSO.FileExists(R_COUNT_REMEMBER_ID) THEN
					Fso.CopyFile Array_FileName(R_COUNT_REMEMBER_ID), Array_FileName(RI)
				END IF
			END IF
		END IF
		' IF R_COUNT=R_COUNT_REMEMBER THEN 
			' IF Array_PSTART_COUNTER(RI)<>Array_PSTART_COUNTER(R_COUNT_REMEMBER) THEN 
				' IF Array_FileName(R_COUNT_REMEMBER_ID)<> Array_FileName(RI) THEN 
					' Fso.CopyFile Array_FileName(R_COUNT_REMEMBER_ID), Array_FileName(RI)
				' END IF
			' END IF
		' END IF
	END IF
NEXT

' Wscript.Echo "HH"

WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------

Sub READ_PSTART_FOR_BIGGER()
    
    Dim SET_GO
    
	FOR RI=1 TO UBOUND(Array_FileName)

	'\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\PStart\PStart.xml
	
		FILEN_1=Array_FileName(RI)
		IF FSO.FileExists(FILEN_1) THEN
			Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
			Do Until File_1.AtEndOfStream
				NET_PATH_1 = File_1.ReadLine+VBCRLF
				
				SET_GO = FALSE
				If InStr(NET_PATH_1, "<execution>") > 0 Then SET_GO = TRUE
				If SET_GO = True Then
					NET_PATH_2 = File_1.ReadLine+VBCRLF
					NET_PATH_3 = File_1.ReadLine+VBCRLF
					NET_PATH_4 = File_1.ReadLine+VBCRLF
					If InStr(NET_PATH_2, "<count>") > 0 Then
						Array_PSTART_COUNTER(RI)=NET_PATH_1+NET_PATH_2+NET_PATH_3+NET_PATH_4
					END IF
					EXIT DO
				End If
			Loop
			File_1.Close
		End If
	NEXT
End Sub


Sub READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY()
    
    Dim SET_GO
    
    ArrayCount = 0

	FILEN_1="C:\NETWORK_COMPUTER_NAME.txt"
	IF FSO.FileExists(FILEN_1) THEN
		Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
		Do Until File_1.AtEndOfStream
			NET_PATH = File_1.ReadLine
			
			SET_GO = True
			If InStr(NET_PATH, "BTHUB") > 0 Then SET_GO = False
			If InStr(NET_PATH, "NAS-QNAP-ML") > 0 Then SET_GO = False
			If SET_GO = True Then
				ArrayCount = ArrayCount + 1
				Array_NETPATH_01(ArrayCount) = NET_PATH
				Array_NETPATH_02(ArrayCount) = Replace(NET_PATH, "-", "_")
				ELEMENT1 = "\\"
				ELEMENT2 = Array_NETPATH_01(ArrayCount)
				ELEMENT3 = "\"
				ELEMENT4 = Array_NETPATH_02(ArrayCount)
				ELEMENT5 = FileName_2
				ELEMENT7 = "_" + NET_PATH

				Array_FileName(ArrayCount) = ELEMENT1 + ELEMENT2 + ELEMENT3 + ELEMENT4 + ELEMENT5' + ELEMENT7
			End If
		Loop
		File_1.Close
		
		ReDim Preserve Array_FileName(ArrayCount)
		ReDim Preserve Array_NETPATH_01(ArrayCount)
		ReDim Preserve Array_NETPATH_02(ArrayCount)
		ReDim Preserve Array_PSTART_COUNTER(ArrayCount)
	END IF
End Sub




Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun
