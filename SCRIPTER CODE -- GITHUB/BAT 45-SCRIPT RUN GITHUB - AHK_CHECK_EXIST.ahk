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

; #Persistent
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

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

LOOP, %0%  ; FOR EACH PARAMETER:
  INFO .= %A_Index% " "
  
COMMAND_PARAM=%INFO%
IF !COMMAND_PARAM
	RETURN
	; NONE COMMAND LINE INFO ASK TO SEARCH FOR
	
FINDER=0

WINGET, ID, list,ahk_class ConsoleWindowClass
LOOP, %ID%
{
	TABLE := ID%A_Index%
	WinGetTitle, Title, ahk_id %TABLE%

	IF INSTR(Title,COMMAND_PARAM)
		FINDER+=1
} 

EXITAPP, %FINDER%
; -------------------------------------------------------------------
; EXIT CODE RETURN HOW MANY FINDER COUNT
; -------------------------------------------------------------------

RETURN
; -------------------------------------------------------------------
 

 
;# ------------------------------------------------------------------
; END BLOCK OF CODE -- EXIT ROUTINE
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
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
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
; EXIT THE APP
; -------------------------------------------------------------------

