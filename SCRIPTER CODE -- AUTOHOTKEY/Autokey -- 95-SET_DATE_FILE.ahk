;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
;# __ 
;# __ Autokey -- 94-ALL LOWER THAN NORMAL PROCCES PRIORITY RESTORE.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# Fri 28-Feb-2020 22:44:00 -- 1 HOUR 54 MINUTE -- GOT GO AR
;# Fri 28-Feb-2020 23:58:00 -- 3 HOUR 08 MINUTE -- ADD MORE
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Fri 28-Feb-2020 23:58:00 -- 3 HOUR 08 MINUTE -- ADD MORE
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
; Sat 29-Feb-2020 15:10:00 -- 1 HOUR 13 MINUTE
; Text Size  16303 -- CRC32 53AC75E0
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
		SUBST_1:= SubStr(A_LoopReadLine, 1, 11)
		SUBST_2:= SubStr(A_LoopReadLine, 15)
		; TOOLTIP % SUBST_1 "`n" SUBST_2,100,100
		WORK_DO_COUNT+=1
	}

	IF WORK_DO_COUNT
		MSGBOX, 4096,, % "CHANGE DATE COUNTER =`n" WORK_DO_COUNT
	
RETURN

