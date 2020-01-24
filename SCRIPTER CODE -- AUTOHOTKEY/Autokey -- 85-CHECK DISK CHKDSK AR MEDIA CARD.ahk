;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 85-CHECK DISK CHKDSK AR MEDIA CARD.ahk
;# __ 
;# __ Autokey -- 85-CHECK DISK CHKDSK AR MEDIA CARD.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Sun 19-Jan-2020 08:05:24 10:20:00
;# UNTIL       TIME [ Sun 19-Jan-2020 10:20:00 -- 2 HOUR 15 MINUTE
;# __ 

;# LAST EDITOR TIME [ Tue 21-Jan-2020 16:35:01
;# UNTIL       TIME [ Tue 21-Jan-2020 21:00:00 -- 4 HOUR 23 MINUTE
;# __ 
;  =============================================================

; ------------------------------------------------------------------
; Location Internet
;---------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL_2.ahk
; ----
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; I SPENT 
; -------------------------------------------------------------------
; Sun 19-Jan-2020 08:05:24
; Sun 19-Jan-2020 10:20:00 -- 2 HOUR
; -------------------------------------------------------------------
; TO LEARN THE FORMAT FOR PUT CODE ONE LINER
; REASON MY BATCH FILE HAS CMD TO EXECUTE RUN
; AND WHEN IT DONE PUT TITLE OF WINDOW FOR THE CMD %comspec%
; BUTT BASTARD PROBLEM
; IF TITLE IS IN THE BATCH CODE
; IT WILL TAKE A TIME BEFORE COME IN
; AND PREVIOUS IT WILL HAVE 
; C:\Windows\System32\cmd.exe
; SO UNABLE TO SEARCH FINDWINDOW LOCATER
; NOT FOR FEW VITAL SECOND LOST
; -------------------------------------------------------------------
; FIRST TIME I CODE THIS UP
; AND BEFORE MANY TIME WANTER
; NONE MY CODE INCLUDE THIS SNIPPET BEFORE
; -------------------------------------------------------------------
; BETWEEN VISUAL BASIC AND _.BAT AND _.VBS VBCRIPT
; VERY HARD SEARCH INTERNET
; NONE EXAMPLE CONSIDER THE PROBLEM BEHAVIOUR OF 
; TITLE WANT IMMEDIATELY - WHY ISSUE DELAY
; THERE NOT A LOT TO IT WHEN LOOKER
; AS LONG QUOTATION MARK PROPER PLACE
; AS LONG GOT THE CODE NOW
; -------------------------------------------------------------------
; SEEING AS AUTOHOTKEYS CAME LATER 
; THAN VISUAL BASIC AND _.BAT AND _.VBS VBSCRIPT
; THAT THE DELAY FOR
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; COMMAND LINE EXAMPLE FOR USER WITH GOODSYNC
; -------------------------------------------------------------------
; "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 85-CHECK DISK CHKDSK AR MEDIA CARD.ahk" NOT_SHOW
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 003
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; I SPENT BEFORE SESSION 002
; -------------------------------------------------------------------
; Sun 19-Jan-2020 08:05:24
; Sun 19-Jan-2020 10:20:00 -- 2 HOUR
; -------------------------------------------------------------------
; I SPENT ANOTHER -- TOO FINISH OFF COMPLETE
; -------------------------------------------------------------------
; Tue 21-Jan-2020 16:35:01
; Tue 21-Jan-2020 16:35:01 -- 3 HOUR 50 MINUTE
; -------------------------------------------------------------------
; DON'T MISTAKE TO OBSERVE
; CODE HERE IS TWIN WITH MY TIMER VERSION RUN EVERY 24 HOUR
; AND PICKUP FOR REMOTE COMPUTER WANT RUN
; -------------------------------------------------------------------
; HERE PROJECT CODE SET FILENAME AND ROUTINE HEADER
; -------------------------------------------------------------------
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE_TIMER:
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; -------------------------------------------------------------------
; ALSO FIND NEW COMPUTER CODER TECHNO TIP
; WHEN RUN FEW LINE OF BATCH FILE _.BAT
; AND PURPOSE TO DELETE AFTER
; SEEM ERROR ALWAYS THROW ERROR NOT TRAPPABLE ABOUT 
; LAST LINE OF NATCH FILE GONE
; IN THAT CASE WRITE THE BATCH AS ALL ONE LINE WITH JOIN 
; COMMAND SYMBOL & OR && WILL STOP RUN MORE IF ERROR BEFORE
; & ALSO FIND 
; NOTEPAD++ ABLE USE DOCUMENT SWITCHER
; RIGHT CLICK FILE DOCUMENT THERE
; SELECT MOVE TO NEW VIEW
; PICK A FEW PROJECT AND WORK THEM SMALL GROUP
; RATHER THAN LOAD OF PROJECT -- WAS LOOK OTHER DAY FOR SEARCHER BE HELP
; PLUGIN HAS SEARCH -- BUT FAIL THAT ONE
; ALSO EASIER SORT OF DOC SWITCH PANEL
; FOUND IN WINDOW PULL DOWN MENU NOT SEEM WORK GOOD AND 
; NONE HOTKEY THAT BASTARD -- NOT WORK SECOND TIME
; ALSO OTHER DAY FIND 
; BUTT CRASHES 
; SourceCookifier v0.7.3.0
; GOOD SHOW PANEL OF ALL ROUTINE IN SCRIPT AND OTHER 
; BUTT ISSUE NOT LIKE LARGE FILE OF TEXT WHEN NOT CODE
; -------------------------------------------------------------------
; Sun 19-Jan-2020 08:05:24
; Sun 19-Jan-2020 10:20:00 -- 2 HOUR
; -------------------------------------------------------------------
; TO LEARN THE FORMAT FOR PUT CODE ONE LINER
; REASON MY BATCH FILE HAS CMD TO EXECUTE RUN
; AND WHEN IT DONE PUT TITLE OF WINDOW FOR THE CMD %comspec%
; BUTT BASTARD PROBLEM
; IF TITLE IS IN THE BATCH CODE
; IT WILL TAKE A TIME BEFORE COME IN
; AND PREVIOUS IT WILL HAVE 
; C:\Windows\System32\cmd.exe
; SO UNABLE TO SEARCH FINDWINDOW LOCATER
; NOT FOR FEW VITAL SECOND LOST
; -------------------------------------------------------------------
; FIRST TIME I CODE THIS UP
; AND BEFORE MANY TIME WANTER
; NONE MY CODE INCLUDE THIS SNIPPET BEFORE
; -------------------------------------------------------------------
; BETWEEN VISUAL BASIC AND _.BAT AND _.VBS VBCRIPT
; VERY HARD SEARCH INTERNET
; NONE EXAMPLE CONSIDER THE PROBLEM BEHAVIOUR OF 
; TITLE WANT IMMEDIATELY - WHY ISSUE DELAY
; THERE NOT A LOT TO IT WHEN LOOKER
; AS LONG QUOTATION MARK PROPER PLACE
; AS LONG GOT THE CODE NOW
; -------------------------------------------------------------------
; SEEING AS AUTOHOTKEYS CAME LATER 
; THAN VISUAL BASIC AND _.BAT AND _.VBS VBSCRIPT
; THAT THE DELAY FOR
; -------------------------------------------------------------------			
; CMD  -- /C OR /K -- /C WILL EXIT OUT -- BATCH FILE REQUIRE EXIT COMMAND ENDER
; START - /B WILL RUN IN SAME DOS SHELL PROMPT WINDOW -- ALL IN ONES
; -------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
#NoTrayIcon

; -------------------------------------------------------------------
; #Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
; OnExit("ExitFunc")

; Register an object to be called on exit:
; OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01-INCLUDE MENU 01 of 03.ahk
;# ------------------------------------------------------------------
; SCRIPT ============================================================

; DetectHiddenWindows, on
; SetStoreCapslockMode, off

; SoundBeep , 2000 , 100
; SoundBeep , 2500 , 100


OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6
; IF OSVER_N_VAR=10
; 	WANT_GO=TRUE

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
WANT_GO=
IF (A_ComputerName="7-ASUS-GL522VW")
	WANT_GO=TRUE
IF (A_ComputerName="4-ASUS-GL522VW")
	WANT_GO=TRUE

	
IF WANT_GO
	GOSUB RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
	
RETURN

; -------------------------------------------------------------------
; CODE IS RUN TIMER FROM HERE 
; -------------------------------------------------------------------
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; -------------------------------------------------------------------
; AND ROUTINE
; -------------------------------------------------------------------
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; -------------------------------------------------------------------

RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:

	; ---------------------------------------------------------------
	; INSTRUCTION FOR BATCH FILE SET
	; ---------------------------------------------------------------
	
	SCRIPT_DIR__:=SubStr(A_ScriptDir, 3)       ; FROM 3 ONWARD INCLUDE 1ST \
	StringReplace, SCRIPT_DIR__,SCRIPT_DIR__,\SCRIPTER\,\SCRIPTOR DATA\, All
	
	SCRIPT_NAME_VAR_CHKDSK_1=%SCRIPT_DIR__%\Autokey -- 19-SCRIPT_TIMER_UTIL_2_BAT_#NFS.BAT

	XY_1=%A_COMPUTERNAME%
	StringReplace, XY_2,XY_1,-,_, All
	; ---------------------------------------------------------------
	; SEMI EXAMPLE -- StringReplace -- 7-ASUS-GL522VW,7_ASUS_GL522VW
	; ---------------------------------------------------------------
	
	SCRIPT_NAME_VAR_CHKDSK_2=\\%XY_1%\%XY_2%_01_C_DRIVE%SCRIPT_DIR__%\Autokey -- 19-SCRIPT_TIMER_UTIL_2_BAT_#NFS.BAT
	IF INSTR(XY_1,"4-")>0 
		StringReplace, SCRIPT_NAME_VAR_CHKDSK_3,SCRIPT_NAME_VAR_CHKDSK_2,4,7, All
	IF INSTR(XY_1,"7-")>0 
		StringReplace, SCRIPT_NAME_VAR_CHKDSK_3,SCRIPT_NAME_VAR_CHKDSK_2,7,4, All
	
	; ---------------------------------------------------------------
	; %SCRIPT_NAME_VAR_CHKDSK_2% WILL BE LOCAL COMPUTER FIRST VIA NETWORK PATH
	; ---------------------------------------------------------------
	
	; ---------------------------------------------------------------
	; CHKDSK.EXE V: /F/X -- X FORCE DISMOUNT OR QUESTION PROMPT
	; ---------------------------------------------------------------
	MIDNIGHT_CHKDSK_MC_V:="@ECHO OFF`n@CD\`n@TITLE Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk`n@ECHO.`n@ECHO.`n@VOL V:`n@ECHO.`n@ECHO.`nCHKDSK.EXE V: /F/X`nEXIT`n" ; `nPAUSE`n"
	MIDNIGHT_CHKDSK_MC_V:="@ECHO OFF&@CD\&@TITLE Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk&@ECHO.&@ECHO.&@VOL V:&@ECHO.&@ECHO.&CHKDSK.EXE V: /F/X&EXIT`n" ; `nPAUSE`n"

	IFEXIST, X:\MP_ROOT\
	StringReplace, MIDNIGHT_CHKDSK_MC_V,MIDNIGHT_CHKDSK_MC_V,&CHKDSK.EXE V: /F/X,&CHKDSK.EXE V: /F/X&@ECHO.&@ECHO.&CHKDSK.EXE X: /F/X, All
	
	; NONE FORWARD AND BACKSLASH -- AOT ATOS COUNCIL

	; ---------------------------------------------------------------
	; NOW REQUIRE WRITE ALL COMMAND OF ONE LINE
	; SAVE DEBUG DO FAULT END IF DELETE AFTER BEGIN
	; ---------------------------------------------------------------
	
	; ---------------------------------------------------------------
	; REMOTE FIRST THEN LOCAL
	; WILL ACT SOON AS FIND BY TIMER ANOTHER CODE PROJECT _.AHK
	; ---------------------------------------------------------------
	
	IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_3%
	{
		FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_3%
		FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_3%
	}
	
	IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_2%
	{
		FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_2%
		FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_2%
	}
	
	; ---------------------------------------------------------------
	; Command-line arguments - Rosetta Code
	; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
	; ---------------------------------------------------------------
	; AutoHotkey
	; From the AutoHotkey documentation: "The script sees incoming parameters as the variables %1%, %2%, 
	; and so on. In addition, %0% contains the number of parameters passed (0 if none). "
	; ---------------------------------------------------------------
	Command_Params=
	Loop %0% ; number of parameters
		Command_Params .= %A_Index% . A_Space
	IF %0%=0
		Command_Params=
	; ---------------------------------------------------------------

	SetTitleMatchMode 2  ; Specify PARTIAL path
	DetectHiddenWindows, on
	
	; WRITE TO NETWORK PATH AND BOTH EXIST LOCAL
	; ---------------------------------------------------------------
	
	IFEXIST, %SCRIPT_NAME_VAR_CHKDSK_1%
	IFWINNOTEXIST Autokey -- 19-SCRIPT_TIMER_UTIL_2 ahk_class ConsoleWindowClass
	{
		IF INSTR(Command_Params,"NOT_SHOW")>0
		{
			IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_1%
			{
				FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%
				FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_1%
			}
			RunWAIT, %comspec% /C "START /B /WAIT "Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk" "%SCRIPT_NAME_VAR_CHKDSK_1%"",,HIDE
			FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%
		}
		
		; ---- NOT_SHOW")=0 IS TRUE_SHOW ----------------------------
		IF INSTR(Command_Params,"NOT_SHOW")=0  ;  ----
		{
			; -------------------------------------------------------
			; I SPENT 
			; -------------------------------------------------------
			; Sun 19-Jan-2020 08:05:24
			; Sun 19-Jan-2020 10:20:00 -- 2 HOUR
			; -------------------------------------------------------
			; I SPENT ANOTHER -- TOO FINISH OFF COMPLETE
			; -------------------------------------------------------
			; Tue 21-Jan-2020 16:35:01
			; Tue 21-Jan-2020 16:35:01 -- 3 HOUR 50 MINUTE
			; -------------------------------------------------------
			; DON'T MISTAKE TO OBSERVE
			; CODE HERE IS TWIN WITH MY TIMER VERSION RUN EVERY 24 HOUR
			; AND PICKUP FOR REMOTE COMPUTER WANT RUN
			; -------------------------------------------------------
			; HERE PROJECT CODE SET FILENAME AND ROUTINE HEADER
			; -------------------------------------------------------
			; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
			; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE_TIMER:
			; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
			; -------------------------------------------------------
			; ALSO FIND NEW COMPUTER CODER TECHNO TIP
			; WHEN RUN FEW LINE OF BATCH FILE _.BAT
			; AND PURPOSE TO DELETE AFTER
			; SEEM ERROR ALWAYS THROW ERROR NOT TRAPPABLE ABOUT 
			; LAST LINE OF NATCH FILE GONE
			; IN THAT CASE WRITE THE BATCH AS ALL ONE LINE WITH JOIN 
			; COMMAND SYMBOL & OR && WILL STOP RUN MORE IF ERROR BEFORE
			; & ALSO FIND 
			; NOTEPAD++ ABLE USE DOCUMENT SWITCHER
			; RIGHT CLICK FILE DOCUMENT THERE
			; SELECT MOVE TO NEW VIEW
			; PICK A FEW PROJECT AND WORK THEM SMALL GROUP
			; RATHER THAN LOAD OF PROJECT -- WAS LOOK OTHER DAY FOR SEARCHER BE HELP
			; PLUGIN HAS SEARCH -- BUT FAIL THAT ONE
			; ALSO EASIER SORT OF DOC SWITCH PANEL
			; FOUND IN WINDOW PULL DOWN MENU NOT SEEM WORK GOOD AND 
			; NONE HOTKEY THAT BASTARD -- NOT WORK SECOND TIME
			; ALSO OTHER DAY FIND 
			; BUTT CRASHES 
			; SourceCookifier v0.7.3.0
			; GOOD SHOW PANEL OF ALL ROUTINE IN SCRIPT AND OTHER 
			; BUTT ISSUE NOT LIKE LARGE FILE OF TEXT WHEN NOT CODE
			; -------------------------------------------------------
			; Sun 19-Jan-2020 08:05:24
			; Sun 19-Jan-2020 10:20:00 -- 2 HOUR
			; -------------------------------------------------------
			; TO LEARN THE FORMAT FOR PUT CODE ONE LINER
			; REASON MY BATCH FILE HAS CMD TO EXECUTE RUN
			; AND WHEN IT DONE PUT TITLE OF WINDOW FOR THE CMD %comspec%
			; BUTT BASTARD PROBLEM
			; IF TITLE IS IN THE BATCH CODE
			; IT WILL TAKE A TIME BEFORE COME IN
			; AND PREVIOUS IT WILL HAVE 
			; C:\Windows\System32\cmd.exe
			; SO UNABLE TO SEARCH FINDWINDOW LOCATER
			; NOT FOR FEW VITAL SECOND LOST
			; -------------------------------------------------------
			; FIRST TIME I CODE THIS UP
			; AND BEFORE MANY TIME WANTER
			; NONE MY CODE INCLUDE THIS SNIPPET BEFORE
			; -------------------------------------------------------
			; BETWEEN VISUAL BASIC AND _.BAT AND _.VBS VBCRIPT
			; VERY HARD SEARCH INTERNET
			; NONE EXAMPLE CONSIDER THE PROBLEM BEHAVIOUR OF 
			; TITLE WANT IMMEDIATELY - WHY ISSUE DELAY
			; THERE NOT A LOT TO IT WHEN LOOKER
			; AS LONG QUOTATION MARK PROPER PLACE
			; AS LONG GOT THE CODE NOW
			; -------------------------------------------------------
			; SEEING AS AUTOHOTKEYS CAME LATER 
			; THAN VISUAL BASIC AND _.BAT AND _.VBS VBSCRIPT
			; THAT THE DELAY FOR
			; -------------------------------------------------------			
			; CMD  -- /C OR /K -- /C WILL EXIT OUT -- BATCH FILE REQUIRE EXIT COMMAND ENDER
			; START - /B WILL RUN IN SAME DOS SHELL PROMPT WINDOW -- ALL IN ONES
			; -------------------------------------------------------
			IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_1%
			{
				FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%
				FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_1%
			}

			RunWAIT, %comspec% /C "START /B /WAIT "Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk" "%SCRIPT_NAME_VAR_CHKDSK_1%""
			IFEXIST, %SCRIPT_NAME_VAR_CHKDSK_1%
				FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%

			; -------------------------------------------------------
			; NOT HAVE RunWAIT -- WILL DEL LINE CREATE SO QUICK NOT RUN AT ALL
			; -------------------------------------------------------

			; -------------------------------------------------------
			; STILL MAYBE PROBLEM OF REMOTE COMPUTER NOT FINISH CHKDSK YET 
			; BEFORE GOODSYNC START JOB
			; -------------------------------------------------------
			
			
			; -------------------------------------------------------
			; 001 AND 002
			; -------------------------------------------------------
			; EXAMPLE VB LINE CODE
			; -------------------------------------------------------
			; Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus
			; -------------------------------------------------------
			; 002
			; -------------------------------------------------------
			; EXAMPLE AHK CODE -- %COMSPEC% ONLY AND PIPE _ > _ COMMAND AFTER
			; -------------------------------------------------------
			; Run %comspec% /c ""%FN_VAR%" "/ALL" "" >"%A_TEMP%\IPTEST.TXT , , MIN
			; -------------------------------------------------------
		}
	}

	
RETURN

