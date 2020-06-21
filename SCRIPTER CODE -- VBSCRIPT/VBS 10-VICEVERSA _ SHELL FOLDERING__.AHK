;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK
;# __ 
;# __ VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# START TIME [ Fri 22:19:40 Pm_08 Jun 2018 ]
;# END   TIME [ Sat 00:10:00 Pm_09 Jun 2018 ] 1 HOUR 40 MINUTE
;# __ 
;====================================================================

; ------------------------------------------------------------------- 
; DESCRIPTION 
; ------------------------------------------------------------------- 
; THE SORT MY C:\ ROOT DIRECTORY OUT KEEPING IT IN SYNC TO A 
; SUB-FOLDER OFF OF C:\ DRIVE
; ------------------------------------------------------------------- 

; ------------------------------------------------------------------- 
; ------------------------------------------------------------------- 
; Location OnLine
; ---------------
; Link to Folder of My AutoHotKeys Project Set
; https://drive.google.com/open?id=0BwoB_cPOibCPVmVYT1pKWUk4LVE

; Link to Folder of My AutoHotKeys Project Set Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0

; Link to This File On DropBox With Most Up to Date
; 
; ------------------------------------------------------------------- 

; ------------------------------------------------------------------- 
; 002
; ------------------------------------------------------------------- 
; ADD CARRIED FORWARD ARGUMENTS FOR GUI MODE VS QUITE
; ------------------------------------------------------------------- 
; FROM  ---- Thu 21-Jun-2018 00:51:49
; TO    ---- Thu 21-Jun-2018 01:02:00
; ------------------------------------------------------------------- 

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

SET_GO=TRUE
Process, Exist, wscript.exe
If ErrorLevel > 0 ; errorlevel will = 0 if process doesn't exist
{
	SET_GO=FALSE 
}
Process, Exist, ViceVersa.exe
If  ErrorLevel > 0 ; errorlevel will = 0 if process doesn't exist
{	
	SET_GO=FALSE 
}

String=

NoP = %0%
If (NoP > 0) 
{
	Loop, %NoP% {
		Par := %A_Index%
		String := String . Par . ", "
	}
	StringTrimRight, String, String, 2
}

IF SET_GO=TRUE 
{
	;  MSGBOX "VICE VERSA AUTOHOTKEYS " %STRING%
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 10-VICEVERSA _ SHELL FOLDERING__.VBS"
	IfExist, %FN_VAR%
		RunWait, "%FN_VAR%" %String%
}
