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
#singleInstance force
detectHiddenWindows, on
setWorkingDir %a_scriptDir%
; #NoTrayIcon

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

RENAME_FILTER_EXTENSION_CASE_EXCLUDE:="MP4 MPG"
RENAME_FILTER_EXTENSION_CASE_EXCLUDE=
GOSUB TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER

EXITAPP

RETURN


#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 84-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk


