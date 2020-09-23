;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 72-RUN HUBIC WITH DELAY.ahk
;# __ 
;# __ Autokey -- 72-RUN HUBIC WITH DELAY.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; BIT HEAVY ON SYSTEM AT BOOTER
; NOT A DAY DELAY
; -------------------------------------------------------------------
; FROM __ Mon 04-Mar-2019 08:50:00
; TO   __ Tue 26-Feb-2019 09:10:00
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; ----
; Autokey -- 58-Auto Repeat Browser Function Set.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2058-Auto%20Repeat%20Browser%20Function%20Set.ahk
; ----
;# ------------------------------------------------------------------


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
#Persistent
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


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

	
If (OSVER_N_VAR=5)
{	
	SOUNDBEEP 1000,100
	EXITAPP
	RETURN
}

	
Process, Exist, hubiC.exe
If ErrorLevel ; errorlevel will = 0 if process doesn't exist
{	
	SOUNDBEEP 1000,100
	EXITAPP
	RETURN
}
; ----
; Minutes to Milliseconds Conversion (min to ms)
; https://www.timecalculator.net/minutes-to-milliseconds
; ----
	
IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName="MATT 04")
{
	; SETTIMER, RUN_HUBIC, 3600000 ; 60 min = 3600000 ms
	; SETTIMER, RUN_HUBIC, 2400000 ; 40 min = 2400000 ms
	SETTIMER, RUN_HUBIC, 40000
	RETURN
}

IF (A_ComputerName="5-ASUS-P2520LA" and A_UserName<>"MATT 01")
	EXITAPP
IF (A_ComputerName="4-ASUS-GL522VW" and A_UserName<>"MATT 01")
	EXITAPP
IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName<>"MATT 04")
	EXITAPP
IF (A_ComputerName="8-MSI-GP62M-7RD" and A_UserName<>"MATT 01")
	EXITAPP
	
SETTIMER, RUN_HUBIC, 40000 ; 40 Second


	
RETURN


;--------------------------------------------------------------------
RUN_HUBIC:
;-------------------------------------------
; REQUIRE LITTLE DELAY ON RUN HUBIC
; GETS SOME CLASH RUN FROM REGISTRY ALSO RUN
;-------------------------------------------

Process, Exist, hubiC.exe
If ErrorLevel ; errorlevel will = 0 if process doesn't exist
{
	EXITAPP
	RETURN
}
	
; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5)
	{
		Process, Exist, hubiC.exe
		If Not ErrorLevel ; errorlevel will = 0 if process doesn't exist
		{
			IfWinNotExist , hubiC
			{
				FN_VAR:="C:\Program Files\OVH\hubiC\hubiC.exe"
				IfExist, %FN_VAR%
				{
					Run, "%FN_VAR%"
					
					RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, hubiC
					
					; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 40-RUN EXE.VBS" "%FN_VAR%"
					SoundBeep , 2500 , 100
					EXITAPP
				}
			}
		}
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
; exit the app

