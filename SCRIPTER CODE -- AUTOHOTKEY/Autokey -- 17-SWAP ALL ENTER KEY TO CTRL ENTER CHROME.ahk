;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 17-SWAP ALL ENTER TO CTRL ENTER WEB PAGES.ahk
;# __ 
;# __ Autokey -- 17-SWAP ALL ENTER TO CTRL ENTER WEB PAGES.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ Wed 18 October 2017 02:05:00-------
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; 
;# ------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------
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

;# ----------------------------------------------------------------------------
;# ----------------------------------------------------------------------------

; SCRIPT ======================================================================

DetectHiddenWindows, on

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100
WinNotActive_Idle = 0
SUSPEND
PAUSE
SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

setTimer TIMER_SUB_1,1000
setTimer TIMER_PREVIOUS_INSTANCE,4000


TIMER_SUB_1:
	;SoundBeep , 2500 , 100

	;ifWinNotActive ahk_class Chrome_WidgetWin_1
	IfWinNotActive ahk_group GroupNameTest_1
	{
		WinNotActive_Idle += 1
		;SoundBeep , 3500 , 100
	} 
	else 
	{
		WinNotActive_Idle = 0
	}

	if (WinNotActive_Idle > 40)
	{
		SoundBeep , 2500 , 100
		WinNotActive_Idle = 0
		SUSPEND
		PAUSE
		SoundBeep , 2500 , 100
	}
Return

	
;Ctrl enter unable to be swapped back to Enter so use SendInput
;--------------------------------------------------------------
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_2
;^enter::
;sendinput {ENTER}
;SetCapsLockState , Off
;SoundBeep , 2500 , 100
;Return

;}	






#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


TIMER_PREVIOUS_INSTANCE:

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





;# ----------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ----------------------------------------------------------------------

;# ----------------------------------------------------------------------
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
;# ------------------------------------------------------------------------------
; exit the app



;Tooltip

; sets title matching to search for "containing" instead of "exact"
;SetTitleMatchMode, 2
	
;Tooltip % ALLOW_VAR

;If ALLOW_VAR = "True" 
;^enter::b
	
;}

;If WinActive("ahk_class Chrome_WidgetWin_1") ALLOW := "True"

;if WinActive("ahk_class Chrome_WidgetWin_1")
;{
;	if WinActive("A" mysms)
;	ALLOW := "False"
;}
;#IfWinactive ahk_class Chrome_WidgetWin_1 
;#IfWinActive mysms ALLOW := "False"
;#IfWinActive mysms ALLOW := "False"


; ----------------------------------------------------------------
; With group here mysms will be in the exclude
; If chrome exist but not if another window in chrome
; ----------------------------------------------------------------
;GroupAdd, GroupNameTest_1, ahk_class Chrome_WidgetWin_1, , , mysms
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_1
;enter::^enter
;return
	
;Ctrl enter unable to be swapped back to Enter so use SendInput 
;#IfWinActive ahk_group GroupNameTest_2
;^enter::sendinput {CapsLock}{ENTER}
;^enter::
;return


;ALLOW_VAR := "False"
;If WinActive("ahk_class" "Chrome_WidgetWin_1") 
;{
;ALLOW_VAR := "True"
;}
;MSGBOX % ALLOW_VAR
;if WinActive("mysms") 
;{
;ALLOW_VAR := "False"
;}
;MSGBOX % ALLOW_VAR

;#IfWinactive ahk_class Chrome_WidgetWin_1 
;enter::
;#IfWinActive mysms 
;^enter::SoundBeep , 2500 , 100

;If ALLOW_VAR = "True" 
;{
;sendinput, {ENTER}
;}

;RETURN

