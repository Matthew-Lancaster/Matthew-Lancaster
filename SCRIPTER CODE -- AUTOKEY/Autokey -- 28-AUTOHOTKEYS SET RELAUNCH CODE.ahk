;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
;# __ 
;# __ Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Sat 02-Mar-2019 21:44:00
; TO   TIME __ Sat 02-Mar-2019 22:34:00
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk· GitHub
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2028-AUTOHOTKEYS%20SET%20RELOADER.ahk
; ----
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.

SetStoreCapslockMode, off
DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

; SETTIMER TIMER_PREVIOUS_INSTANCE,1

SETTIMER TIMER_SUB_AUTOHOTKEY_RELOAD,1000

RETURN


TIMER_SUB_AUTOHOTKEY_RELOAD:

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
TEMP_VAR_1=%FN_VAR_1%
TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")

FN_VAR_1=%TEMP_VAR_3%

IFWINEXIST %FN_VAR_1%
{
	FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
	AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
	TEMP_VAR_1=%FN_VAR_1%
	TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
	TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
	TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
	FN_VAR_2=%TEMP_VAR_3%

	WinGet, PID, PID, %FN_VAR_2% ahk_class AutoHotkey
	Process, Close,% PID
	RETURN
}

IFWINNOTEXIST %FN_VAR_1%
{
	SoundBeep , 2000 , 100
	Run, %FN_VAR_1%
	EXITAPP
}

RETURN




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
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT
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

