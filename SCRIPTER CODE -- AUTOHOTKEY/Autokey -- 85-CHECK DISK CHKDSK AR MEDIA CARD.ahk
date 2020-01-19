;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# __ Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Sun 28-Apr-2019 14:38:04 ]
;# LAST EDIT TIME [ Tue 20-Aug-2019 15:40:00 ]
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
; AS LONG QUOTATION MARK RIGHT PLACE
; AS LONG GOT THE CODE NOW
; -------------------------------------------------------------------
; SEEING AS AUTOHOTKEYS CAME LATER 
; THAN VISUAL BASIC AND _.BAT AND _.VBS VBSCRIPT
; THAT THE DELAY FOR
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

; ---------------------------------------------------------------
; ---------------------------------------------------------------

; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01-INCLUDE MENU 01 of 03.ahk

;# ------------------------------------------------------------------

; SCRIPT ============================================================


DetectHiddenWindows, on
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
IF (A_ComputerName<>"7-ASUS-GL522VW")
	WANT_GO=TRUE
IF (A_ComputerName<>"4-ASUS-GL522VW")
	WANT_GO=TRUE

MIDNIGHT_CHKDSK_MEDIA_CARD_V_HDD=
OLD_MIDNIGHT_CHKDSK_MEDIA_CARD_V_HDD=

MIDNIGHT_CHKDSK_MEDIA_CARD_V_HDD= %A_Now%
	
IF WANT_GO=TRUE
	GOSUB RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
	
RETURN


; -------------------------------------------------------------------
; CODE IS RUN TIMER FROM HERE
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:

	; ---------------------------------------------------------------
	; NOT REALLY REQUIRE TO TWO SET STRING FILENAME
	; USE FIX VARIABLE HARD CODE AND NOT PICK A_ScriptName
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	; INSTRUCTION FOR BATCH FILE SET
	; ---------------------------------------------------------------
	SCRIPT_NAME_VAR_CHKDSK_1=%A_ScriptDir%\Autokey -- 19-SCRIPT_TIMER_UTIL_2_BAT_#NFS.BAT
	
	SCRIPT_DIR__:=SubStr(A_ScriptDir, 3)                          ; FROM 3 ONWARD
	SCRIPT_NAME_VAR_CHKDSK_2=\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE%SCRIPT_DIR__%\Autokey -- 19-SCRIPT_TIMER_UTIL_2_BAT_#NFS.BAT

	
	StringReplace, SCRIPT_NAME_VAR_CHKDSK_3,SCRIPT_NAME_VAR_CHKDSK_2,7-ASUS-GL522VW,4-ASUS-GL522VW, All
	StringReplace, SCRIPT_NAME_VAR_CHKDSK_3,SCRIPT_NAME_VAR_CHKDSK_2,7_ASUS_GL522VW,4_ASUS_GL522VW, All
	
	; ---------------------------------------------------------------
	; CHKDSK.EXE V: /F/X -- X FORCE DISMOUNT OR QUESTION PROMPT
	; ---------------------------------------------------------------
	MIDNIGHT_CHKDSK_MC_V:="@ECHO OFF`n@CD\`n@TITLE Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk`n@ECHO.`n@ECHO.`n@VOL V:`n@ECHO.`n@ECHO.`nCHKDSK.EXE V: /F/X`nEXIT`n"  ; `nPAUSE`n"
	; ---------------------------------------------------------------

	FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%
	FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_1%
	IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_2%
	{
		FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_2%
		FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_2%
	; 	MSGBOX %SCRIPT_NAME_VAR_CHKDSK_2%
	}
	IFNOTEXIST, %SCRIPT_NAME_VAR_CHKDSK_3%
	{
		FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_3%
		FileAppend,%MIDNIGHT_CHKDSK_MC_V%,%SCRIPT_NAME_VAR_CHKDSK_3%
	}
	
	; ----
	; Command-line arguments - Rosetta Code
	; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
	; ----
	; AutoHotkey
	; From the AutoHotkey documentation: "The script sees incoming parameters as the variables %1%, %2%, 
	; and so on. In addition, %0% contains the number of parameters passed (0 if none). "

	Command_Params=
	
	Loop %0% ; number of parameters
		Command_Params .= %A_Index% . A_Space
	IF %0%=0
		Command_Params=

	SetTitleMatchMode 2  ; Specify PARTIAL path

	IFEXIST, %SCRIPT_NAME_VAR_CHKDSK_1%
	IFWINNOTEXIST Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk ahk_class ConsoleWindowClass
	{
		IF INSTR(Command_Params,"NOT_SHOW")>0
			RunWAIT, %SCRIPT_NAME_VAR_CHKDSK_1% ,,HIDE
		
		IF INSTR(Command_Params,"NOT_SHOW")=0
		{
			
			; -------------------------------------------------------
			; I SPENT 
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
			; AS LONG QUOTATION MARK RIGHT PLACE
			; AS LONG GOT THE CODE NOW
			; -------------------------------------------------------
			; SEEING AS AUTOHOTKEYS CAME LATER 
			; THAN VISUAL BASIC AND _.BAT AND _.VBS VBSCRIPT
			; THAT THE DELAY FOR
			; -------------------------------------------------------

			; -------------------------------------------------------
			; CMD  -- /C OR /K -- /C WILL EXIT OUT -- BATCH FILE REQUIRE EXIT COMMAND ENDER
			; START - /B WILL RUN IN SAME DOS SHELL PROMPT WINDOW -- ALL IN ONES
			; -------------------------------------------------------
			
			RunWAIT, %comspec% /C "START /B "Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk" "%SCRIPT_NAME_VAR_CHKDSK_1%""
			
			; -------------------------------------------------------
			; EXAMPLE VB LINE CODE
			; -------------------------------------------------------
			; Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus
			RETURN
			; -------------------------------------------------------
			
			; -------------------------------------------------------
			; EXAMPLE AHK CODE -- %COMSPEC% ONLY AND PIPE _ > _ COMMAND AFTER
			; -------------------------------------------------------
			; Run %comspec% /c ""%FN_VAR%" "/ALL" "" >"%A_TEMP%\IPTEST.TXT , , MIN
			; -------------------------------------------------------
		}
	}

	FileDELETE, %SCRIPT_NAME_VAR_CHKDSK_1%
	
RETURN

