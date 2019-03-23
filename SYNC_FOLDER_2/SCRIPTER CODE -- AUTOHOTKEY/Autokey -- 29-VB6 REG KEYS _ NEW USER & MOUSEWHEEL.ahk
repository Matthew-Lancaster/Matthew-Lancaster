;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 29-VB6 REG KEYS _ NEW USER.ahk
;# __ 
;# __ Autokey -- 29-VB6 REG KEYS _ NEW USER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Wed 13:07:30 Pm_09 May 2018 ]
;# END         TIME [ Wed 14:10:00 Pm_09 May 2018 ]
;# LAST EDITOR TIME [ Wed 14:10:00 Pm_09 May 2018 ]
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; 001 ---------------------------------------------------------------
; THIS WILL RUN AND SET VB 6 IF NOT BEEN SET FIRST TIME
; BY CHECKING CERTAIN VALUE IN REGISTRY
; AND ALSO PUT SOME RECENT FILE IN FOR EDITOR
; AND ALSO SORT THE MOUSEWHEEL REGISTRERS IT AND ALSO CHECKS 
; THE CHECKBOXES REQUIRED
; -------------------------------------------------------------------
; FROM TIME __ Wed 09-May-2018 13:07:51 
; TO   TIME __ Wed 09-May-2018 14:10:00 1 HOUR+
; -------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
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


DetectHiddenWindows, on

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

global SET_GO
global OSVER_N_VAR

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

; -------------------------------------------------------------------
GOSUB MAIN_ROUTINE

SOUNDBEEP 1000,200
SOUNDBEEP 1500,200

Exitapp

RETURN
; -------------------------------------------------------------------

MAIN_ROUTINE:

	Process, Exist, VB6.EXE
	If ErrorLevel >0 
		RETURN

	GOSUB VB_SET_MOUSE_WHEEL
	GOSUB VB_SET_MOUSE_WHEEL_CHECK
	GOSUB VB_SET_RECENT
	GOSUB VB_SET_REG_KEY_IF_WANTING

RETURN
; -------------------------------------------------------------------


VB_SET_MOUSE_WHEEL:

;Reg.exe add "HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect" /v "LoadBehavior" /t REG_DWORD /d "3" /f

	Count_Reg_MouseWheel=0

	Loop, Reg, HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		;MSGBOX , %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%, %value%
		Count_Reg_MouseWheel+=1

	}

	; CHECK THIS WANTING
	; THEN MUST MEAN SETTINGS NEVER BEEN SET BEFORE
	If Count_Reg_MouseWheel=0
	{
		SOUNDBEEP 2000,200
		FN_VAR:="D:\VB6\VB-NT\#__ MOUSE_WHEEL_FIX\VB6IDEMouseWheelAddin __ RUN ME.BAT"
		IfExist, %FN_VAR%
			Run, "%FN_VAR%"
	}


RETURN

VB_SET_MOUSE_WHEEL_CHECK:

	Loop, Reg, HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		;MSGBOX , %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%, %value%

		; CHECK THIS WANTING
		; THEN MUST MEAN SETTINGS NEVER BEEN SET BEFORE
		IfInString, A_LoopRegName, LoadBehavior
		If value<>3
		{
			SOUNDBEEP 2000,200
			FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\REG KEYS __ VB 6\VB 6 _ REG KEYS MOUSE WHEEL.BAT"
			IfExist, %FN_VAR%
				Run, "%FN_VAR%"
			BREAK
		}
	}


RETURN



VB_SET_REG_KEY_IF_WANTING:

	Loop, Reg, HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		;MSGBOX , %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%, %value%
		
		
		; CHECK THIS NAME IF NOT SET TO VALUE WANTING
		; THEN MUST MEAN SETTINGS NEVER BEEN SET BEFORE
		IfInString, A_LoopRegName, SaveBeforeRun
		If value<>2
		{
			SOUNDBEEP 2000,200
			FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\REG KEYS __ VB 6\VB 6 _ REG KEYS.BAT"
			IfExist, %FN_VAR%
				Run, "%FN_VAR%" , , MIN
			Break
		}
	}


RETURN


VB_SET_RECENT:
	Count_Reg_Recent=0

	Loop, Reg, HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\RecentFiles\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		Count_Reg_Recent+=1
		;MSGBOX , %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%
		
		
		;IfInString, value, Google\Chrome\Application
		;IfInString, value, \Installer\chrmstp.exe
		;{
		;RegDelete, %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%
		;MsgBox, 4, , %A_LoopRegKey%\%A_LoopRegSubKey%
		;MsgBox, 4, , %A_LoopRegName% = %value% (%A_LoopRegType%)`n`nContinue?
		;IfMsgBox, NO, break
		;}
	}
	IF Count_Reg_Recent<3 
	{
		SOUNDBEEP 2000,200
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\REG KEYS __ VB 6\VB 6 _ REG KEYS RECENT.BAT"
		IfExist, %FN_VAR%
			Run, "%FN_VAR%" , , MIN
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
;# -----------------------------------------------------------------0
; exit the app


