;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 51-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
;# __ 
;# __ Autokey -- 51-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Sun 05:28:00 Am_17 Feb 2019 ]
;# LAST EDIT TIME [ Sun 05:55:00 Am_17 Feb 2019 ]
;# __ 
;  =============================================================


; -------------------------------------------------------------------
; 001
; -------------------------------------------------------------------
; FROM TO Sun 17-Feb-2019 05:55:00
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Stop the Tunisia of Advert Arrive Our Way
; It Already Begun On Windows 7 Pro Computer Intel 5
; The UORIGIN BlOCK Show It Blocking Naught Item While Advert Are Appeared There Every Where
; I Shall Put this Block in Own File To Make Sure Run is All the Time
; And Also Up the PACE of TIMER So UPDATE By GOOGLE Is Stopped Instantly
; Rather Than Having Partial Download Corrupting Anything
; Not That was Skillful Window 7 Machine Has Finish Some Type of Update 
; Involving Extensions Now Is Okay Advert Free
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location Internet
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 54-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2054-Google%20Chrome%20Update%20Process%20Killer%20Stop%20the%20Tunisia%20of%20Advert.ahk
; ----
;# ------------------------------------------------------------------

#SingleInstance Force

IF A_ComputerName=1-ASUS-X5DIJ
	ExitApp     
IF A_ComputerName=2-ASUS-EEE
	ExitApp     

ExitApp     

	
#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk



;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

; ROBOFORM MYSMS FILL AND SUBMIT
; #InstallKeybdHook ;<- add this
; #InstallMouseHook ;<- this
; #UseHook, On ;<- and this


; -------------------------------------------------------------------
; LEFT RUN LONG TIMER AND OUT OF DATE WAS THIS MUCH
; CNET PUBLISH STORY OF AMERICAN BULL TROLLOP AND THEN 
; HOW AN ADVERT INVASION WAS ABOUT TO TAKE OVER
; PLAGGY FROM MY LEARN 
; TALK ABOUT ADVERT SAME TIME
; NOTHING CAME ABOUT OF THAT THING
; TOO HARD NOW-A-DAY IF THINK LEARN
; STOP THIS BIT AND THAT
; SURPRISE FOR ADVERT PEOPLE
; [ Sunday 12:33:40 Pm_26 May 2019 ]
; THIS NOT A MAIN CODE LEFT RUN ANYMORE FOR ME
; QUITE EFFECTIVE AT KILL UPDATE CHROME
; -------------------------------------------------------------------
; Google Chrome (64bit)
; Publisher: Google
; License: Freeware
; Installed Version: 74.0.3729.131
; -------------------------------------------------------------------



; DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

SETTIMER TIMER_PREVIOUS_INSTANCE,1

SETTIMER TIMER_KILL_GOOGLE_CHROME_UPDATE_GOING_TO_USE_AD_BLOCK_KILLER ,1000

TIMER_KILL_GOOGLE_CHROME_UPDATE_GOING_TO_USE_AD_BLOCK_KILLER:

SET_GO=FALSE
Process, Exist, GoogleUpdate.exe
If ErrorLevel
	SET_GO=TRUE

; MSGBOX % SET_GO
	
IF SET_GO=TRUE
{
	; Process, Close, GoogleUpdate.exe
	Run, "TASKKILL.exe" /F /IM GoogleUpdate.exe /T , , HIDE
	; SoundBeep , 2000 , 100
}
	
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
; EXIT THE APP
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

