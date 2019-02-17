;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 51-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
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
; Location OnLine
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 54-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2054-Google%20Chrome%20Update%20Process%20Killer%20Stop%20the%20Tunisia%20of%20Advert.ahk
; ----
;# ------------------------------------------------------------------


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


;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

; ROBOFORM MYSMS FILL AND SUBMIT
; #InstallKeybdHook ;<- add this
; #InstallMouseHook ;<- this
; #UseHook, On ;<- and this

; DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

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

