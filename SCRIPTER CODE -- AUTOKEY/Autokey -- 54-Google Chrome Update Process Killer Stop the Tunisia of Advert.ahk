;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 51-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
;# __ 
;# __ Autokey -- 51-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Sun 05:28:00 Am_17 Feb 2019 ]
;# LAST EDIT TIME [ Sun 05:55:00 Am_17 Feb 2019 ]
;# __ 
;  =============================================================


; -------------------------------------------------------------------
; 001
; -------------------------------------------------------------------
; FROM TO Sun 17-Feb-2019 05:55:00
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Stop the Tunisia of Advert Arrive Our Way
; It Already Begun On Windows 7 Pro Computer Intel 5
; The UORIGIN BlOCK Show It Blocking Naught Item While Advert Are Appeared There Every Where
; I Shall Put this Block in Own File To Make Sure Run is All the Time
; And Also Up the PACE of TIMER So UPDATE By GOOGLE Is Stopped Instantly
; Rather Than Having Partial Download Corrupting Anything
; Not That was Skillful Window 7 Machine Has Finish Some Type of Update 
; Involving Extensions Now Is Okay Advert Free
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location OnLine
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 54-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2054-Google%20Chrome%20Update%20Process%20Killer%20Stop%20the%20Tunisia%20of%20Advert.ahk
; ----
;# ------------------------------------------------------------------

#SingleInstance Force

IF A_ComputerName=1-ASUS-X5DIJ
	ExitApp     
IF A_ComputerName=2-ASUS-EEE
	ExitApp     

	
#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------
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


;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

; ROBOFORM MYSMS FILL AND SUBMIT
; #InstallKeybdHook ;<- add this
; #InstallMouseHook ;<- this
; #UseHook, On ;<- and this

; DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

SETTIMER TIMER_PREVIOUS_INSTANCE,1

SETTIMER TIMER_KILL_GOOGLE_CHROME_UPDATE_GOING_TO_USE_AD_BLOCK_KILLER ,1000

TIMER_KILL_GOOGLE_CHROME_UPDATE_GOING_TO_USE_AD_BLOCK_KILLER:

SET_GO=FALSE
Process, Exist, GoogleUpdate.exe
If ErrorLevel
	SET_GO=TRUE

; MSGBOX % SET_GO
	
IF SET_GO=TRUE
{
	; Process, Close, GoogleUpdate.exe
	Run, "TASKKILL.exe" /F /IM GoogleUpdate.exe /T , , HIDE
	; SoundBeep , 2000 , 100
}
	
RETURN


MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process,Close,% DllCall("GetCurrentProcessId")
	
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
		; Process,Close, AutoHotKey.exe
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
; EXIT THE APP
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

