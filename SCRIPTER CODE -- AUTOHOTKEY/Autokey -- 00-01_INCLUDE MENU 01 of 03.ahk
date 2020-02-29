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
; TO   __ Sun 09-Jun-2019 10:00:00 __ Clipboard Count = 139 __ NEAR 3 HOUR
; AND THEN 
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------

; -------------------------------------------------
; I COMBINE 01 INTO 02 
; IT MAKE EASIER ADD ENTRY
; AND BIG GOSUB ROUTINE ARE REMAIN NUMBER THREE #03
; -------------------------------------------------
; Tue 18-Feb-2020 19:54:00
; -------------------------------------------------

; -------------------------------------------------------------------
; FIRST TIME USER GOTO ON AUTOHOTKEYS
; -------------------------------------------------------------------
; GOTO NIPER_MenuHandler
; -------------------------------------------------------------------
; UNAVOID
; OR MOVE ALL MY EDITOR BACK AGAIN
; HERE #INCLUDE THE MENU AT EARLY PART OF CODER
; AND THEN ROUTINE GO LATER WHEN MOST PROJECT GET ROUTINE DONE
; AFTER EDITOR NEW CHANGE MENU AT EARLY FRONT NOW LINK WITH 
; PART TWO SMALL ROUTINE TAKE MENU GOING
; NOT EASY 
; OR EDIT BACK WAY BEFORE 
; WITH MENU ITEM NEAR END OF CODE
; -------------------------------------------------------------------
; Tue 18-Feb-2020 20:34:07
; -------------------------------------------------------------------

Menu, Tray, Add  ; Creates a separator line.

FileGetTime, A_Script_MODIIFED_DATE , A_ScriptFullPath, M
FormatTime, TimeString, A_Script_MODIIFED_DATE, yyyy MMM dd hh:mm:ss tt
Menu, Tray, Add, %A_ScriptName% ---- %TimeString%, MenuHandler  ; Creates a new menu item.

; Create the popup menu by adding some items to it.
if A_ScriptName=Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
{
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, RELAUNCH CODE, MenuHandler  ; Creates a new menu item.
}

if A_ScriptName=Autokey -- 21-AUTORUN.ahk
{
VAR_RUN_ME_NOW_AUTOBOOT=FALSE
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, RUN HERE NOW, MenuHandler  ; Creates a new menu item.
}


Menu, Tray, Add  
Menu, Tray, Add, Terminate Script, MenuHandler
Menu, Tray, Add, Terminate All AutoHotKey.exe -- LEFT(Ctrl)+WInKEY+ESCAPE, MenuHandler
Menu, Tray, Add  
Menu, Tray, Add, RESTORE_VB_KEEP_RUNNER AND ELITESPY -- RIGHT(Ctrl)+F1, MenuHandler
Menu, Tray, Add  
Menu, Tray, Add, RELOAD    ALL NET - VB CODE.exe, MenuHandler 
Menu, Tray, Add, TERMINATE ALL NET - VB CODE.exe, MenuHandler 
Menu, Tray, Add  
Menu, Tray, Add, RELOAD    All AutoHotKey Network.exe, MenuHandler  
Menu, Tray, Add, TERMINATE All AutoHotKey Network.exe, MenuHandler  
Menu, Tray, Add  

if A_ScriptName=Autokey -- 72-RUN HUBIC WITH DELAY.ahk
{
Menu, Tray, Add  
Menu, Tray, Add, 1 HOUR TIMER UNTIL HUBIC.EXE LEARN, MenuHandler  
}


if A_ScriptName=Autokey -- 58-Auto Repeat Browser Function Set.ahk
{
Menu, Tray, Add
Menu, Tray, Add, Pause __ Debby Hall, MenuHandler 
}

; if A_ScriptName=Autokey -- 14-BRIGHTNESS WITH DIMMER.ahk
; {
Menu, Tray, Add
Menu, Tray, Add, ADD 1 HOUR BEFORE SCREEN SAVER, MenuHandler 
; }

Menu, Tray, Add, CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK

Menu, Tray, Add, TIMEZONE_MINI_GUI_DISPLAY_EXE

Menu, Tray, Add, ALL_LOW_PROCCES_PRIORITY_TO_NORMAL

Menu, Tray, Add  


; -------------------------------------------------------------------
; FIRST TIME USER GOTO ON AUTOHOTKEYS
; UNAVOID
; OR MOVE ALL MY EDITOR BACK AGAIN
; HERE #INCLUDE THE MENU AT EARLY PART OF CODER
; AND THEN ROUTINE GO LATER WHEN MOST PROJECT GET ROUTINE DONE
; AFTER EDITOR NEW CHANGE MENU AT EARLY FRONT NOW LINK WITH 
; PART TWO SMALL ROUTINE TAKE MENU GOING
; NOT EASY 
; OR EDIT BACK WAY BEFORE 
; WITH MENU ITEM NEAR END OF CODE
; -------------------------------------------------------------------
; Tue 18-Feb-2020 20:34:07
; -------------------------------------------------------------------
; ADJUST AGAIN AS ERROR TRY USE LABEL AND GOTO
; NOW USER IF THEN WITH TRUE=FALSE SO NEVER RUN THROUGH
; BUTT HOPE WHEN CALL THE ROUTINE PICK UP DO
; MAYBE REQUIRE EXIT WITH BREAK OR CONTINUE TO GET OFF IT
; -------------------------------------------------------------------
; Tue 18-Feb-2020 21:13:02
; -------------------------------------------------------------------
; GOTO NIPER_MENUHANDLER
; -------------------------------------------------------------------

IF TRUE=FALSE
{
MenuHandler:

; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.

; A_ThisMenuItem=RELOAD ALL NET - VB CODE.exe
; A_ThisMenuItem=KILL   ALL NET - VB CODE.exe
; TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE

if A_ThisMenuItem=RELAUNCH CODE
{
	Run, "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk"
	Process, Close,% DllCall("GetCurrentProcessId")
}

if A_ThisMenuItem=RUN HERE NOW
{
	VAR_RUN_ME_NOW_AUTOBOOT=TRUE

}


if A_ThisMenuItem=RESTORE_VB_KEEP_RUNNER AND ELITESPY -- RIGHT(Ctrl)+F1
{
	GOSUB SUB_RESTORE_VB_KEEP_RUNNER
	GOSUB SUB_RESTORE_ELITESPY
}

if A_ThisMenuItem=Terminate Script
	Process, Close,% DllCall("GetCurrentProcessId")

if A_ThisMenuItem=Terminate All AutoHotKey.exe -- LEFT(Ctrl)+WInKEY+ESCAPE
{
	; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
	GOSUB TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME
	
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

; 01 OF 04
; ---------------------------------------------------------------
; ---------------------------------------------------------------
if A_ThisMenuItem=RELOAD    ALL NET - VB CODE.exe
{
	GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_01_OF_04
}

; 02 OF 04
; ---------------------------------------------------------------
; ---------------------------------------------------------------
if A_ThisMenuItem=TERMINATE ALL NET - VB CODE.exe
{
	GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_02_OF_04
}

; 03 OF 04
; ---------------------------------------------------------------
; ---------------------------------------------------------------
if A_ThisMenuItem=RELOAD    All AutoHotKey Network.exe
{
	GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_03_OF_04
}

; 04 OF 04
; ---------------------------------------------------------------
; ---------------------------------------------------------------
if A_ThisMenuItem=TERMINATE All AutoHotKey Network.exe
{
	GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_04_OF_04
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

if A_ThisMenuItem=ADD 1 HOUR BEFORE SCREEN SAVER
{
	ADD_MINUTE_BEFORE_SCREEN_SAVER=TRUE
	GOSUB WRITE_FILE_SCREEN_BRIGHT_FOR_1_HOUR
}

if A_ThisMenuItem=CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK
{
	GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK
}

if A_ThisMenuItem=TIMEZONE_MINI_GUI_DISPLAY_EXE
{
	GOSUB TIMEZONE_MINI_GUI_DISPLAY_EXE
}


if A_ThisMenuItem=ALL_LOW_PROCCES_PRIORITY_TO_NORMAL
{
	GOSUB ALL_LOW_PROCCES_PRIORITY_TO_NORMAL
}


; -------------------------------------------------------------------
; UNABLE TO MAKER CODE AFTER HERE
; ALL CODE ROUTINE GO IN 
; Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk
; -------------------------------------------------------------------

RETURN
}

; NIPER_MenuHandler:
