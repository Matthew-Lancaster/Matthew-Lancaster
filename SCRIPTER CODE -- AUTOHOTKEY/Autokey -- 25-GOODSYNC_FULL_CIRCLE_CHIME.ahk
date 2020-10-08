;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 25-GOODSYNC_FULL_CIRCLE_CHIME.ahk
;# __ 
;# __ Autokey -- 25-GOODSYNC_FULL_CIRCLE_CHIME.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Fri 16:24:00 Am_12 Apr 2017 ]
;# LAST EDIT TIME [ Fri 16:48:00 Pm_27 Apr 2018 ] 24 Minute
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; CODE IS USER WHEN GOODSYNC DO A FULL CYCLE OF SYNC AND 
; BACK TO BEGINNING AGAIN OR JUST STARTED
; SET AT THE 4 HOUR TIME MOST JOBS SET SET IN
; IT IS LOADED FROM A SCRIPT AS THE FIRST JOB
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; ONLINE LOCATE HERE
; ----
; SCRIPTER CODE -- AUTOKEY – Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
; ----
; -------------------------------------------------------------------
; WELL HERE IS THE SHARE FILE
; ----
; Autokey -- 25-GOODSYNC_FULL_CIRCLE_CHIME.ahk
; https://www.dropbox.com/s/puf965euzqeslkp/Autokey%20--%2010-READ%20MOUSE%20CURSOR%20ICON%20STATE%20AND%20BEEPER%20WHEN%20NOT%20BUSY%20HOUR%20GLASS%20OVER.ahk?dl=0
; ----
; -------------------------------------------------------------------

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force
; ---------------------------------------------------------------
; USE HERE #Persistent OR WILL LEAVE THE CODE UPON ONE CYCLE OF A RUNNER
; ---------------------------------------------------------------
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

;---------------------------------------------------------
; CODE INITIALIZE SOUND EFFECT LEARN
;---------------------------------------------------------
SoundBeep , 1500 , 400
;SoundBeep , 2500 , 400

SoundBeep , 3500 , 400
;SoundBeep , 4500 , 400

Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\Autokey -- 25-GOODSYNC_FULL_CIRCLE_CHIME\BBC Micro.wav
SLEEP, 1000
Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\Autokey -- 25-GOODSYNC_FULL_CIRCLE_CHIME\BBC Micro.wav
SLEEP, 1000

Exitapp

return
;--------------------------------------------------------------------
		
		
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
		
; ---------------------------------------------------------------
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
; ---------------------------------------------------------------


