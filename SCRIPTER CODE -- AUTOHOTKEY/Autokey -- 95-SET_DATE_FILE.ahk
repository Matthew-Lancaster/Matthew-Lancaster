;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
;# __ 
;# __ Autokey -- 94-ALL LOWER THAN NORMAL PROCCES PRIORITY RESTORE.ahk
;# __ 
;# BY __ Matt.Lan@Btinternet.com
;# __ 
;# Sun 26-Apr-2020 19:17:03
;# Sun 26-Apr-2020 23:58:00 -- 4 HOUR 40 MINUTE
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; WORK TIME
; -------------------------------------------------------------------
; SESSION 2 DAY 1
; Sun 26-Apr-2020 11:22:41
; Sun 26-Apr-2020 11:05:46 -- 12 HOUR 56 MINUTE
; -------------------------------------------------------------------
; SESSION 2 DAY 1
; Sun 26-Apr-2020 19:17:03
; Sun 26-Apr-2020 23:58:00 -- 4 HOUR 40 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; TO CREATE EFFORT FOR NEW MP3 UNIT
; -------------------------------------------------------------------
; 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; ------------------------------------------------------------------
; Location SOURCE
;---------------------------------------------------------------------
; ---------------------
; Matthew-Lancaster/Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
; ---------------------
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2094-ALL_LOW_PROCCES_PRIORITY_TO_NORMAL.ahk 
; ---------------------
; HTTPS://TINYURL.COM/SRZ69YL
; ---------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CREDIT TO 
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CREDIT TO **** AUTOHOTKEYS **** HELP PAGE -- SEARCH WORD -- Process
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; MAIN CREDIT MY OPERATION
; A MASSIVE EFFORT HARD WORK FIND READY MADE LEARN SCRIPT 
; LIKE HERE
; -------------------------------------------------------------------
; -------------------------------------------------------------------

#Warn
#NoEnv
#SingleInstance Force


RENAME_EXTENSION_SET_GO_EVERY_TIME=
O_LEN_STRING_HOLD_HWND=
RENAME_FILTER_EXTENSION_CASE_EXCLUDE=
RENAME_FILTER_EXTENSION_CASE_INCLUDE=
RENAME_EXTENSION_QUIET_WITH_AUDIO=
RENAME_EXTENSION_SET_DONE_QUIET=

GOSUB TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER
	
RETURN

TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER:

	WORK_DO_COUNT=0
	
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	; "F:\MP3-YX-510_02_TS\M"
	Loop,read,C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510_FILE_SCRIPT.TXT
	{
		
		SUBST_2_FILENAME:= SubStr(A_LoopReadLine, 13)

		; DATE NOT ABLE SET BELOW 1600 YEAR
	
		SUBST_1_DATE:= SubStr(A_LoopReadLine, 1, 8)
		SUBST_1_DATE_Y:= SubStr(A_LoopReadLine, 1, 4)
		SUBST_1_DATE_M:= SubStr(A_LoopReadLine, 5,2)
		SUBST_1_DATE_D:= SubStr(A_LoopReadLine, 7,2)

		; SUBST_1_DATE:= StrReplace(SUBST_1_DATE,"/","")

		; FormatTime, TS, %SUBST_1_DATE%, YYYYMMDD
		; FormatTime, TimeString, 20050423220133, dddd MMMM d, yyyy hh:mm:ss tt
		
 		TS:=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00
		; FormatTime, TS, TS, YYYYMMDD

		; IF Mod(A_INDEX, 10)=0 
			; TOOLTIP % SUBST_1_DATE "`n" TS "`n" SUBST_2_FILENAME,100,100
		WORK_DO_COUNT+=1
		
		FileSetTime, TS , %SUBST_2_FILENAME%, M
		
		FileGetTime, TS_2, %SUBST_2_FILENAME%
		IF Mod(A_INDEX, 1000)=0 
			TOOLTIP % SUBST_1_DATE "`n" TS_2 "`n" SUBST_2_FILENAME,100,100
		
	}

	IF WORK_DO_COUNT
		; MSGBOX, 4096,, % "CHANGE DATE COUNTER =`n" WORK_DO_COUNT
	
RETURN

