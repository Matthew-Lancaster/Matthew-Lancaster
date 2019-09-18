;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 84-RENAME_FILE_EXTENSION_CASE_01_VBP.ahk
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
; DECLARER FOR #INCLUDE
; -------------------------------------------------------------------
O_LEN_STRING_HOLD_HWND=
RENAME_EXTENSION_SET_GO_EVERY_TIME=
RENAME_EXTENSION_QUIET_WITH_AUDIO=
RENAME_EXTENSION_SET_DONE_QUIET=
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

RENAME_EXTENSION_SET_GO_EVERY_TIME=TRUE

; RENAME_FILTER_EXTENSION_CASE_EXCLUDE:="MP4 MPG"
RENAME_FILTER_EXTENSION_CASE_INCLUDE:="VBP"

IF (A_ComputerName="1-ASUS-X5DIJ")
	RENAME_EXTENSION_SET_DONE_QUIET=TRUE
IF (A_ComputerName="2-ASUS-EEE")
	RENAME_EXTENSION_SET_DONE_QUIET=TRUE
IF (A_ComputerName="3-LINDA-PC")
	RENAME_EXTENSION_SET_DONE_QUIET=TRUE
IF (A_ComputerName="5-ASUS-P2520LA")
	RENAME_EXTENSION_SET_DONE_QUIET=TRUE

IF (A_ComputerName="4-ASUS-GL522VW")
	RENAME_EXTENSION_QUIET_WITH_AUDIO=TRUE
IF (A_ComputerName="7-ASUS-GL522VW")
	RENAME_EXTENSION_QUIET_WITH_AUDIO=TRUE
IF (A_ComputerName="8-MSI-GP62M-7RD")
	RENAME_EXTENSION_QUIET_WITH_AUDIO=TRUE

	
	
IF RENAME_EXTENSION_SET_DONE_QUIET=TRUE
	RENAME_EXTENSION_QUIET_WITH_AUDIO=TRUE
	
GOSUB TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER

EXITAPP

RETURN


#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 84-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk


