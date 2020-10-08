;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
;# __ 
;# __ Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Sat 02-Mar-2019 21:44:00
; TO   TIME __ Sat 02-Mar-2019 22:34:00
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk· GitHub
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2028-AUTOHOTKEYS%20SET%20RELOADER.ahk
; ----
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp HAVE AR CALL ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; ---------------------------------------------------------------
; I MADE MENU ITEM INTO INCLUDE FILE IN 3 PART 
; 01. INTRO SETUP MENU
; 02. THE MENU ROUTINE
; 03. ANY ROUTINE THE MENU USE
; ---------------------------------------------------------------
; SAVER OF RSI INJURY AND MORE ACCURATE
; THE INCLUDE FILE ARE SAME FOLDER
; ---------------------------------------------------------------
; FROM __ Sun 09-Jun-2019 07:03:00 __ Clipboard Count = 024
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------
; Create the popup menu by adding some items to it.
; MenuHandler:
; ---------------------------------------------------------------
; #Include GO WITH FULL PATH AS SOME LAUNCHER DO NOT SET WORK PATH WHEN RUNNER
; RATHER THAN CHANGE THE WORKING PATH WITHIN-AH
; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk


SetStoreCapslockMode, off
DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

; SETTIMER TIMER_PREVIOUS_INSTANCE,1

SETTIMER TIMER_SUB_AUTOHOTKEY_RELOAD_02,1000
; TIMER_SUB_AUTOHOTKEY_RELOAD__ IS IN THE INCLUDER _ Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


RETURN


TIMER_SUB_AUTOHOTKEY_RELOAD_02:

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
TEMP_VAR_1=%FN_VAR_1%
TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")

FN_VAR_2=%TEMP_VAR_3%

WINEXIST_28_AUTOHOTKEYS_SET_RELOADER=TRUE
WHILE WINEXIST_28_AUTOHOTKEYS_SET_RELOADER=TRUE
{
	IFWINEXIST %FN_VAR_2%
	{
		WINEXIST_28_AUTOHOTKEYS_SET_RELOADER=TRUE
		WinGet, PID_01, PID, %FN_VAR_2% ahk_class AutoHotkey
		SoundBeep , 8000 , 100
		Process, Close,% PID_01
	}
}

; IFWINNOTEXIST %FN_VAR_2%
; {
	SoundBeep , 5000 , 100
	Run, %TEMP_VAR_1%
; }

EXITAPP

RETURN








#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
	return (match > 1)
	}
Return

;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
;# ------------------------------------------------------------------
; exit the app

