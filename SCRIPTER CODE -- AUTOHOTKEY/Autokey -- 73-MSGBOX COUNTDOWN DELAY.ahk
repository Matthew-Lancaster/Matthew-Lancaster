	 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY.ahk
;# __ 
;# __ Autokey -- 73-MSGBOX COUNTDOWN DELAY.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; CODE FROM SEPARATION ANOTHER MAIN CODE
; MAIN CODER WAS WORK ON A LOT SO MOVE HERE
; -------------------------------------------------------------------
; FROM __ Sat 16-Mar-2019 07:37:08
; TO   __ Sat 16-Mar-2019 08:44:00
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location On-Line
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  
Menu, Tray, Add  ; Creates a separator line.

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, oFF
SetTitleMatchMode 2  ; Partial Path

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6


Secs_MSGBOX_01=18
Secs_MSGBOX_02=5
Secs_MSGBOX_03=50
X_COUNT_EXIT=0
MSGBOX_COUNTDOWN_RESTART=""
VAR_WORKER_MSGBOX_DELAY_COUNT=""
OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2
SETTIMER MSGBOX_COUNTDOWN_VB_KEEP_RUNNER_OS_RESTART,1000

SETTIMER MSGBOX_PRESS_FOR_RELOADER,4000
	
	
SETTIMER TIMER_HOTKEY,1000
	
RETURN


TIMER_HOTKEY:

	VAR_IN_NAME=Confirm Save As ahk_class #32770
	SetTitleMatchMode 3  ; Specify Full path
	IFWINEXIST %VAR_IN_NAME%
	IFWINEXIST Confirm Save As ahk_exe VB6.EXE
	{
		ControlGetText CONTROL_TEXT,Button1,%VAR_IN_NAME%
		
		STRING_V:=&&Yes  0
		IF INSTR(CONTROL_TEXT,%STRING_V%)>1
		{	
			; NA [v1.0.45+]: May improve reliability. See reliability below.
			ControlClick, Button1,%VAR_IN_NAME%,,,, NA x10 y10 
			SOUNDBEEP 4000,300
			VAR_DONE_ESCAPE_KEY=TRUE
		}
		IF CONTROL_TEXT=&Yes
		{
			Secs_MSGBOX_04=5
			SOUNDBEEP 5000,200
		}

		IF Secs_MSGBOX_04>0 	
			Secs_MSGBOX_04-=1
			
		ControlSetText,Button1,&Yes  %Secs_MSGBOX_04%, %VAR_IN_NAME%
	}
	
	
	DetectHiddenWindows, ON
	SetTitleMatchMode 2
	IFWINEXIST Microsoft Visual Basic [design] ahk_class wndclass_desked_gsk
	IFWINEXIST [design] - [Object Browser] ahk_class wndclass_desked_gsk
	IFWINEXIST Microsoft Visual Basic [design] ahk_exe VB6.EXE
	{
		; ---------------------------------------------------------------------
		; IN VB6 PRESS F2 BY MISTAKE INSTEAD OF JUMP BACK TO WHERE YOU WERE HOTKEY WITH CONTROL KEY
		; AND BRING UP OBJECT BROWSER AND RIGHT CLICK ON IT HIDE YOU HAVE TO DO
		; SOLOTUION HERE
		; ---------------------------------------------------------------------
		ControlGet, OutputVar_4, Visible, , Object Browser, Microsoft Visual Basic ahk_class wndclass_desked_gsk
		IF OutputVar_4=1
			Control, HIDE,, Object Browser, Microsoft Visual Basic ahk_class wndclass_desked_gsk
		SOUNDBEEP 5000,200
	}

RETURN



MSGBOX_PRESS_FOR_RELOADER:
SET_GO=FALSE
WIN_FIND=Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk ahk_class #32770
IFWINEXIST %WIN_FIND%
	SET_GO=TRUE
WIN_FIND=Autokey -- 32-BRUTE BOOT DOWN.ahk ahk_class #32770
IFWINEXIST %WIN_FIND%
	SET_GO=TRUE
IF SET_GO=TRUE 
{
	ControlGettext, MSGBOX_CONTROL, Static1,%WIN_FIND%
	SET_GO=FALSE
	
	; STRING_SEARCH_VAR=Error:
	; IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		; SET_GO=TRUE
	; STRING_SEARCH_VAR=Warning:
	; IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		; SET_GO=TRUE
	STRING_SEARCH_VAR=Could not close the previous instance of
	IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		SET_GO=TRUE

	IF SET_GO=TRUE
	{
		; MSGBOX % MSGBOX_CONTROL
		; WinActivate, %WIN_FIND%
		; SLEEP 200
		ControlClick,&No, %WIN_FIND%
		ControlClick,Button2, %WIN_FIND%
		ControlClick,Button1, %WIN_FIND%
		SOUNDBEEP 500,50
		WinGet, PID_02, PID, %WIN_FIND%
		Process, Close, PID_02

	}
}
RETURN



MSGBOX_COUNTDOWN_VB_KEEP_RUNNER_OS_RESTART:

VAR_WORKER_MSGBOX_DELAY_COUNT_01=VB KEEP RUNNER ahk_class #32770
VAR_WORKER_MSGBOX_DELAY_COUNT_02=ahk_class #32770 ahk_exe WScript.exe
VAR_WORKER_MSGBOX_DELAY_COUNT_03=Autokey -- ahk_class #32770
VAR_WORKER_MSGBOX_DELAY_COUNT=

IF !VAR_WORKER_MSGBOX_DELAY_COUNT
{
	VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_01%
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
	{
		ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		STRING_SEARCH_COUNTDOWN_DELAY=READY FOR OS RESTART
		IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}
	IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		VAR_WORKER_MSGBOX_DELAY_COUNT=
}

IF !VAR_WORKER_MSGBOX_DELAY_COUNT
{
	VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_02%
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
	{
		ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		STRING_SEARCH_COUNTDOWN_DELAY=GoodSync Script Command to Stop
		IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}
	IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		VAR_WORKER_MSGBOX_DELAY_COUNT=
}

IF !VAR_WORKER_MSGBOX_DELAY_COUNT
{
	VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_03%
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
	{
		ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		STRING_SEARCH_COUNTDOWN_DELAY=Error at Line
		STRING_SEARCH_COUNTDOWN_DELAY=The program will exit
		IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}
	IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		VAR_WORKER_MSGBOX_DELAY_COUNT=
}


IF !VAR_WORKER_MSGBOX_DELAY_COUNT
	RETURN

MSGBOX_ACTIVATE=FALSE

IF OLD_VAR_WORKER_MSGBOX_DELAY_COUNT<>%VAR_WORKER_MSGBOX_DELAY_COUNT%
{
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
	{
		; VB KEEP RUNNER _ REBOOT
		Secs_MSGBOX_01=40
		Secs_MSGBOX_02=5
		MSGBOX_ACTIVATE=TRUE
		X_COUNT_EXIT=0
	}
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
	{
		; GOODSYNC
		Secs_MSGBOX_01=150
		Secs_MSGBOX_02=5
		MSGBOX_ACTIVATE=TRUE
		X_COUNT_EXIT=0
	}
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
	{
		; GOODSYNC
		Secs_MSGBOX_01=24
		Secs_MSGBOX_02=5
		MSGBOX_ACTIVATE=TRUE
		X_COUNT_EXIT=0
	}
}
OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT%

IF %VAR_WORKER_MSGBOX_DELAY_COUNT%
	IF Secs_MSGBOX_01 > -1
	{
		;IF Mod(Secs_MSGBOX_01, 10)=0 
		;	WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		IF MSGBOX_ACTIVATE=TRUE
		{
			MSGBOX_ACTIVATE=FALSE
			WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		}
		
		ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)
		{
			SET_GO=FALSE
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
				SET_GO=TRUE
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
				SET_GO=TRUE
			IF SET_GO=TRUE
			{
				ControlSetText,Static1,%MSGBOX_COUNTDOWN_RESTART%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				ControlSetText,Button1,&Yes  %Secs_MSGBOX_01%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				ControlSetText,Button2,Not, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			}
			SET_GO=FALSE
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
				SET_GO=TRUE
			IF SET_GO=TRUE
			{
				; ControlSetText,Static1,%MSGBOX_COUNTDOWN_RESTART%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				ControlSetText,Button1,&OK  %Secs_MSGBOX_01%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			}
			
			; #WinActivateForce, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			Secs_MSGBOX_01-=1
			RETURN
		}
		ELSE
		{
			X_COUNT_EXIT=0
			RETURN
		}
	}

Secs_MSGBOX_02-=1
	
IF Secs_MSGBOX_01=-1
{
		X_COUNT_EXIT+=1
		IF Mod(X_COUNT_EXIT, 20)=0 
			WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
		IF INSTR(STRING_SEARCH_COUNTDOWN_DELAY,"GoodSync Script Command to Stop")=0
		{
			SET_GO=FALSE
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
				SET_GO=TRUE
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
				SET_GO=TRUE
			; ControlGetText CONTROL_TEXT,Button1,, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			;IF INSTR(CONTROL_TEXT,&Yes  0)=0
			;	SET_GO=FALSE

			IF SET_GO=TRUE
			{
				ControlClick, &Yes  0, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				SOUNDBEEP 500,50
				OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2
			}
		}
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
			; ControlGetText CONTROL_TEXT,Button1,, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			; IF INSTR(CONTROL_TEXT,"OK  0")>0
			{
				ControlClick, OK  0, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				SOUNDBEEP 500,50
				OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2
			}
	
}
Return


SUB_MESS_SPARE_CODE:
; CoordMode, Mouse, SCREEN
		; #WinActivateForce, VB KEEP RUNNER ahk_class #32770
		; WinActivate, VB KEEP RUNNER ahk_class #32770
		; WinGetPos, X_2, Y_2, , , VB KEEP RUNNER ahk_class #32770
		; ControlGetPos, x, y, w, h, Button1, VB KEEP RUNNER ahk_class #32770
		; if Secs_MSGBOX_02>0
			; MouseMove, X+20+X_2, Y+20+Y_2
		
Return







MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		
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

