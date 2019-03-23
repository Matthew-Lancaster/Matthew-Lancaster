;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 72-RUN HUBIC WITH DELAY.ahk
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, 1 HOUR TIMER UNTIL HUBIC.EXE LEARN, MenuHandler  
Menu, Tray, Add  ; Creates a separator line.

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
					
					; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE.VBS" "%FN_VAR%"
					SoundBeep , 2500 , 100
					EXITAPP
				}
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

