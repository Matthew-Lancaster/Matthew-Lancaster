;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 17-SWAP ALL ENTER TO CTRL ENTER WEB PAGES.ahk
;# __ 
;# __ Autokey -- 17-SWAP ALL ENTER TO CTRL ENTER WEB PAGES.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ Wed 18 October 2017 02:05:00-------
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; 
;# ------------------------------------------------------------------


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

;# ----------------------------------------------------------------------------
;# ----------------------------------------------------------------------------

; SCRIPT ======================================================================

DetectHiddenWindows, on

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100
WinNotActive_Idle = 0
SUSPEND
PAUSE
SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

setTimer TIMER_SUB_1,1000
setTimer TIMER_PREVIOUS_INSTANCE,4000


TIMER_SUB_1:
	;SoundBeep , 2500 , 100

	;ifWinNotActive ahk_class Chrome_WidgetWin_1
	IfWinNotActive ahk_group GroupNameTest_1
	{
		WinNotActive_Idle += 1
		;SoundBeep , 3500 , 100
	} 
	else 
	{
		WinNotActive_Idle = 0
	}

	if (WinNotActive_Idle > 40)
	{
		SoundBeep , 2500 , 100
		WinNotActive_Idle = 0
		SUSPEND
		PAUSE
		SoundBeep , 2500 , 100
	}
Return

	
;Ctrl enter unable to be swapped back to Enter so use SendInput
;--------------------------------------------------------------
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_2
;^enter::
;sendinput {ENTER}
;SetCapsLockState , Off
;SoundBeep , 2500 , 100
;Return

;}	
	
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
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			If ( PID_8 <> DllCall("GetCurrentProcessId") ) 
				 ; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				 Process, Close, %PID_8% 
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
		; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe
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


TIMER_PREVIOUS_INSTANCE:

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




;# ----------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ----------------------------------------------------------------------

;# ----------------------------------------------------------------------
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
;# ------------------------------------------------------------------------------
; exit the app



;Tooltip

; sets title matching to search for "containing" instead of "exact"
;SetTitleMatchMode, 2
	
;Tooltip % ALLOW_VAR

;If ALLOW_VAR = "True" 
;^enter::b
	
;}

;If WinActive("ahk_class Chrome_WidgetWin_1") ALLOW := "True"

;if WinActive("ahk_class Chrome_WidgetWin_1")
;{
;	if WinActive("A" mysms)
;	ALLOW := "False"
;}
;#IfWinactive ahk_class Chrome_WidgetWin_1 
;#IfWinActive mysms ALLOW := "False"
;#IfWinActive mysms ALLOW := "False"


; ----------------------------------------------------------------
; With group here mysms will be in the exclude
; If chrome exist but not if another window in chrome
; ----------------------------------------------------------------
;GroupAdd, GroupNameTest_1, ahk_class Chrome_WidgetWin_1, , , mysms
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_1
;enter::^enter
;return
	
;Ctrl enter unable to be swapped back to Enter so use SendInput 
;#IfWinActive ahk_group GroupNameTest_2
;^enter::sendinput {CapsLock}{ENTER}
;^enter::
;return


;ALLOW_VAR := "False"
;If WinActive("ahk_class" "Chrome_WidgetWin_1") 
;{
;ALLOW_VAR := "True"
;}
;MSGBOX % ALLOW_VAR
;if WinActive("mysms") 
;{
;ALLOW_VAR := "False"
;}
;MSGBOX % ALLOW_VAR

;#IfWinactive ahk_class Chrome_WidgetWin_1 
;enter::
;#IfWinActive mysms 
;^enter::SoundBeep , 2500 , 100

;If ALLOW_VAR = "True" 
;{
;sendinput, {ENTER}
;}

;RETURN

