;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 18-NORTON CONTROL BOOTER.ahk
;# __ 
;# __ Autokey -- 18-NORTON CONTROL BOOTER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# 1ST START TIME [ Wed 02:32:20 Am_15 Nov 2017 ]
;# 2ND END   TIME [ Wed 04:14:50 Am_15 Nov 2017 ]
;# __ 
;====================================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
; Location OnLine
; ---------------
; Link to Folder of My AutoHotKeys Project Set
; https://drive.google.com/open?id=0BwoB_cPOibCPVmVYT1pKWUk4LVE

; Link to Folder of My AutoHotKeys Project Set Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0

; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/bsvrrmyknxr7pvd/Autokey%20--%2018-NORTON%20CONTROL%20BOOTER.ahk?dl=0
;# ------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
; --------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
; --------------------
; -------------------------------------------------------------------



settimer GO_SOUND, 2000

RETURN

GO_SOUND:

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

RETURN



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
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return

;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))


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
;# -----------------------------------------------------------------0
; exit the app