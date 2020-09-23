;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 84-RENAME_FILE_EXTENSION_CASE_02.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; SPLIT INCLUDE AH
; -------------------------------------------------------------------
; FROM __ Mon 16-Sep-2019 22:11:33
; TO   __ Mon 16-Sep-2019 22:11:33
; -------------------------------------------------------------------

#noEnv
; #persistent
#SingleInstance force
detectHiddenWindows, on
setWorkingDir %a_scriptDir%
; #NoTrayIcon


; ---------------------------------------------------------------
; #Include GO WITH FULL PATH AS SOME LAUNCHER DO NOT SET WORK PATH WHEN RUNNER
; RATHER THAN CHANGE THE WORKING PATH WITHIN-AH
; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk




; -------------------------------------------------------------------
; DECLARER FOR #INCLUDE
; -------------------------------------------------------------------
O_LEN_STRING_HOLD_HWND=
RENAME_EXTENSION_SET_GO_EVERY_TIME=
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

RENAME_EXTENSION_SET_GO_EVERY_TIME=TRUE

; RENAME_FILTER_EXTENSION_CASE_EXCLUDE:="MP4 MPG VBP"
; RENAME_FILTER_EXTENSION_CASE_INCLUDE:="VBP"

RENAME_EXTENSION_SET_DONE_QUIET=

GOSUB TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER

EXITAPP

RETURN


#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 84-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk







; -------------------------------------------------------------------
; HERE HAS A HOT KEY TERMINATE SMART BOMBER KILL SWITCH FOR AHK
; CONTROL AND WINDOWS-KEY AND F1
; -------------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


