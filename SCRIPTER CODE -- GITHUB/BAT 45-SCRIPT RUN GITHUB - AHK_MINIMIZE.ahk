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

GITHUB_RUNNNER_RERUN=

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

; SoundBeep , 1000 , 200

GITHUB_RUNNNER_RERUN=
BEEN_RUN_DUMMY_RUN_STATUS_BUTTON=FALSE

GOSUB MAIN_ROUTINE

RETURN

; -------------------------------------------------------------------

; GLOBAL SETTINGS ===================================================

; GUI ===============================================================


MAIN_ROUTINE:

Gui, Margin, 5, 5
gui, font, s14 ; , Arial ; , Calibri  
Gui, Add, Button, y+5 w480 gSTATUS, Window of Command Console Minimize

; MINIMIZE_ALL__COMMAND_PROMPT_WITH_GITHUB_ON_REQUEST

Command_Params=

Loop, %0% ; number of parameters
	Command_Params.=%A_Index%
	
IF (!Command_Params or %0%=0)
	Command_Params=GITHUB_RUNNNER


IF GITHUB_RUNNNER_RERUN
	Command_Params=GITHUB_RUNNNER

; Command_Params:=StrReplace(Command_Params, """" , "")
	
EXIT_NOW=TRUE
SOUND_EVENT_DONE=FALSE

WinMinimize, ahk_class Notepad++

WinGet, id, list,ahk_class ConsoleWindowClass
Loop, %id%
{
	Table := id%A_Index%
	WinGetTitle, Title, ahk_id %Table%
	; Command_Params:="%Command_Params%"

	IF INSTR(Title,Command_Params)
	{
		WinGet MMX, MinMax, ahk_id %Table%
		If MMX>-1
		{
			WinMinimize, ahk_id %Table%
			; -----------------------------------------------------------
			; MMX 0 = NORMAL -- MMX 1 = MAXIMIZED -- MMX -1 = MINIMIZED
			; -----------------------------------------------------------
			IF SOUND_EVENT_DONE=FALSE 
			{
				Gui, Show, AutoSize
				SETTIMER TIMER_EXIT, 5000
				SoundBeep , 1000 , 200
				EXIT_NOW=FALSE
				SOUND_EVENT_DONE=TRUE
			}
		}
	}
} 

IF BEEN_RUN_DUMMY_RUN_STATUS_BUTTON=FALSE
	IF INSTR(Command_Params, "QUICK_INTRO_DUMMY_RUN")
	{
		BEEN_RUN_DUMMY_RUN_STATUS_BUTTON=TRUE
		Gui, Show, AutoSize
		EXIT_NOW=FALSE
		SETTIMER TIMER_EXIT, 4000
	}		

IF 	EXIT_NOW=TRUE
	EXITAPP

RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
STATUS:
	GITHUB_RUNNNER_RERUN=GITHUB_RUNNNER
	GOSUB MAIN_ROUTINE
	EXITAPP
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
TIMER_EXIT:
	EXITAPP
RETURN
; -------------------------------------------------------------------

 
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
