;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk
;# __ 
;# __ BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Fri 19-Oct-2018 17:52:00
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

#Persistent
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; --------------------
#SingleInstance force


;# ------------------------------------------------------------------
; DESCRIPTION
;# ------------------------------------------------------------------
; QUITE SIMPLE -- MINIMIZE TWO COMMAND PROMPT WINDOW THAT BEN RUN 
; BY GITHUB CODING INCLUDE THE TWIN GOODSYNC ONE
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; ---- LOCATE ONLINE
; -------------------------------------------------------------------
;
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; ----
; -------------------------------------------------------------------
; FROM   Fri 19-Oct-2018 17:51:00
; TO     Fri 19-Oct-2018 18:10:00
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

SoundBeep , 1000 , 200
; -------------------------------------------------------------------

; GLOBAL SETTINGS ===================================================

; GUI ===============================================================

Gui, Margin, 5, 5
gui, font, s14 ; , Arial ; , Calibri  
Gui, Add, Button, y+5 w480 gSTATUS, Window of Command Console Minimize

; MINIMIZE_ALL__COMMAND_PROMPT_WITH_GITHUB_ON_REQUEST

Command_Params=

Loop %0% ; number of parameters
	Command_Params = %A_Index%

IF !Command_Params
	Command_Params=GIT_RUNNNER

	
WinGet, id, list,ahk_class ConsoleWindowClass
Loop, %id%
{
	Table := id%A_Index%
	WinGetTitle, Title, ahk_id %Table%
	EXIT_NOW=TRUE
	IF INSTR(Title,%Command_Params%)>0
	{
		WinMinimize  ahk_id %Table%
		Gui, Show, AutoSize
		SETTIMER TIMER_EXIT, 4000
		EXIT_NOW=FALSE
	}
} 


IF 	EXIT_NOW=TRUE
	EXITAPP

RETURN

STATUS:
RETURN

TIMER_EXIT:
	EXITAPP
RETURN


 
;# ------------------------------------------------------------------
; USUAL END BLOCK OF CODE TO HELP EXIT ROUTINE
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
; -------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

; -------------------------------------------------------------------
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
; -------------------------------------------------------------------
; exit the app
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; REFERENCE PAGES OPEN 30
; -------------------------------------------------------------------
; -------------------------------------------------------------------
