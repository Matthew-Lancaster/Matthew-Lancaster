;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 75-GOODSYNC OPTIONS SET.ahk
;# __ 
;# __ Autokey -- 75-GOODSYNC OPTIONS SET.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; WORK DO-ER
; -------------------------------------------------------------------
; SPLIT THE CODE AWAY FROM 
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; 01
; AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO:
; 02
; AUTO_RELOAD_RAIN_ALARM:
; 03 
; AUTO_RELOAD_FACEBOOK_QUICK_SUB:
; 04 
; AUTO_RELOAD_FACEBOOK:
; 
; -------------------------------------------------------------------
;
; FROM __ Sun 28-Apr-2019 08:40:38
; TO   __ Sun 28-Apr-2019 08:40:38 __ 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 002 AND 003 
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; ----
; Autokey -- 58-Auto Repeat Browser Function Set.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2058-Auto%20Repeat%20Browser%20Function%20Set.ahk
; ----
;# ------------------------------------------------------------------

;
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

IF A_ComputerName=2-ASUS-EEE
	Exitapp

GLOBAL OSVER_N_VAR

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6


; IF A_ComputerName=1-ASUS-X5DIJ
; IF A_ComputerName=2-ASUS-EEE


SETTIMER SET_OK_BOX,500


RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------

; SET OKAY BOX AFTER MADE SELECTION
SET_OK_BOX:
{

	DetectHiddenWindows, ON
	SetTitleMatchMode 3
	IfWinExist Left Folder ahk_class #32770
		WinActivate, Left Folder ahk_class #32770
		; #WinActivateForce, Left Folder ahk_class #32770
	IfWinActive Left Folder ahk_class #32770
	{	
		ControlGetPos, x, y, , , OK, Left Folder ahk_class #32770
		; MouseMove, X+10, Y+10
		ControlClick, OK, Left Folder ahk_class #32770,,,, NA x20 y20
		SoundBeep , 2000 , 400
	}	
	IfWinExist Right Folder ahk_class #32770
		WinActivate, Right Folder ahk_class #32770
		; #WinActivateForce, Right Folder ahk_class #32770
	IfWinActive Right Folder ahk_class #32770
	{	
		ControlGetPos, x, y, , , OK, Right Folder ahk_class #32770
		; MouseMove, X+10, Y+10
		ControlClick, OK, Right Folder ahk_class #32770,,,, NA x20 y20
		SoundBeep , 3000 , 400
	}	

	
	
	DetectHiddenWindows, ON
	SetTitleMatchMode 2
	IfWinExist Options ahk_class #32770
	WinActivate, Options ahk_class #32770
	;#WinActivateForce, Options ahk_class #32770
	; IfWinExist ] Options ahk_class #32770
	; TOOLTIP OO
	IfWinActive ] Options ahk_class #32770
	{	
		ControlGettext, OutputVar_2, Button16, ] Options ahk_class #32770
		ControlGet, OutputVar_1, Line, 1, Edit9, ] Options ahk_class #32770
		IF (OutputVar_2="Periodically (On Timer), every")
		IF OutputVar_1<>5
		{
			ControlSetText, Edit9,, ] Options ahk_class #32770
			Control, EditPaste, 5, Edit9, ] Options ahk_class #32770
			SoundBeep , 4000 , 100
		}

		IF OutputVar_1=5
		IfWinActive ] Options ahk_class #32770
		{	
			X=0
			ControlGetPos, x, y, , , Button61, Options ahk_class #32770 ; SAVE BUTTON
			; TOOLTIP % X " -- " Y
			IF X>0
			; MouseMove, X+10, Y+10
			IF X>0
			ControlClick, Button61, Options ahk_class #32770,,,, NA x10 y10
			IF X>0
			SoundBeep , 5000 , 400
			X=0
			ControlGetPos, x, y, , , Button63, Options ahk_class #32770 ; SAVE BUTTON
			; TOOLTIP % X " -- " Y
			IF X>0
			; MouseMove, X+10, Y+10
			IF X>0
			ControlClick, Button63, Options ahk_class #32770,,,, NA x10 y10
			IF X>0
			SoundBeep , 5000 , 400
		}	
	}
		
	
}
RETURN



MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		DetectHiddenWindows, On 
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List% 
		  { 
			WinGet, PID, PID, % "ahk_id " List%A_Index% 
			If ( PID <> DllCall("GetCurrentProcessId") ) 
				 ; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				 Process, Close, List%A_Index% 
		  }
		Process, Close,% DllCall("GetCurrentProcessId")
		
		
		;  ----------------------------------------------------------
		; PROBLEM HERE IF PROGRAM THAT CALL THE BATCH FILE IS KILL SO IS THEN BATCH FILE
		; AND WE GET OVER THAT BY GO EXTRA VIA VBSCRIPT ANOTHER FILE
		; COULD OF RUN A  LOOP AND KILL BUT TRY NOT LOSE OWN ONE FIRST
		; [ Saturday 14:55:00 Pm_02 March 2019 ]
		;  ----------------------------------------------------------

		;  ----------------------------------------------------------
		; OTHER OPTION SET PROCESS KILLER
		;  ----------------------------------------------------------
		; Run, BAT_03_PROCESS_KILLER.BAT /F /IM AutoHotKey.exe /T , , Max
		; Run, %ComSpec% /k ""BAT_03_PROCESS_KILLER.BAT" "/F" "/IM" "AutoHotKey.exe" "/T"" , , Max
		; Process, Close, AutoHotKey.exe
		;  ----------------------------------------------------------
	
		; AUTO GENERATED FILE BY HERE VISUAL BASIC ORIGINAL LONG BEFORE AUTOHOTKEY WANT
		; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe
		; D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
		; -------------------------------------------------------------------
		; AND USED BY HERE
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; -------------------------------------------------------------------
		; [ Saturday 14:52:10 Pm_02 March 2019 ]
		; -------------------------------------------------------------------
		; EDITOR COPY PASTE FROM VBS 39-KILL PROCESS.VBS
		; THIS FILE BECAME USE BY
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; AND THEY USE IT HERE THIS ONE
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}

if A_ThisMenuItem=Pause __ Debby Hall
{
	IF DEBBY_HALL_PAUSE=TRUE
	{
		DEBBY_HALL_PAUSE=FALSE
		SOUNDBEEP 5000,200
	}
	ELSE
	{
		DEBBY_HALL_PAUSE=TRUE
		SOUNDBEEP 1000,200
	}
}
return


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

