	;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk
;# __ 
;# __ Autokey -- 32-BRUTE BOOT DOWN.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# START TIME [ Fri 22:19:40 Pm_08 Jun 2018 ]
;# END   TIME [ Sat 00:10:00 Pm_09 Jun 2018 ] 1 HOUR 40 MINUTE
;# EXTRA TIME [ Thu 16:04:00 Pm_29 Nov 2018 ] MAKE ROBUST DOES NOT ALLOW TO BE CLOSED UNLESS FORCE
;# __ 
;====================================================================




;# ------------------------------------------------------------------
; DESCRIPTION 
;# ------------------------------------------------------------------
; ANY PROGRAMS THAT ARE KNOWN TO HAVE BOOT DOWN PROBLEM
; WITH MESSENGER AT SHUTDOWN __ END NOW
; THE BUTTON IS PRESSED FOR THEM AFTER GRACE SHORT TIME
; LEAVING ANY PROGRAM THAT REQUIRE A GRACEFUL SHUT DOWN TO BE LEFT ALONE
; SYSTEMEXPLORER HAS A SHUTDOWN FAULT SOME COMPUTER
; FILLEZILLA SERVER HAS A SHUTDOWN FAULT SOME COMPUTER
;
; GOOGLESYNCDRIVE REQUIRE GRACEFUL SHUTDOWN
; GOODSYNC REQUIRE GRACEFUL SHUT DOWN
;
; THIS PROGRAM WILL NOT GET KNOCKED OUT BY SHUTDOWN REQUEST
; ONLY UNTIL THE OTHER PROGRAMS ARE GONE
; YOU CAN EXIT WITH RIGHT CLICK TO CLOSE 
; ONLY SHUTDOWN REQUEST MODE IS INTERCEPTED
; SHUTDOWN OR LOGG OFF
;
; THIS PROGRAM IS CALLED FROM AT BOOTER ONLY IF REQUIRED 
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AUTORUN.ahk
;
; AT SHUTDOWN WINDOWS REPORT THE WINDOW_TITLE AS PROBLEM PROGRAM
; BUT NOT ALWAYS GIVES AWAY THE .EXE TALKING ABOUT
; USE THIS CODE TO TIDY UP AND BAD SHUTDOWN PROGRAMS
; WITH THE REM LINES LEFT IN TO CONVERT WINDOW_TITLE TO EXE PATH
;
; FINALLY THE PROCESS IS MORE SPEEDY 
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; MAKE ROBUST DOES NOT ALLOW TO BE CLOSED UNLESS FORCE
; ALL CHECKED OVER
;
; MOVE ONEXIT ROUTINE UP MORE SO THEY ALL TO BE LIKE GO 
; THROUGH THE DECLARE STAGE
;
; PUT EXTRA THAT ESCAPE KEY IS PRESS AT SHUTDOWN RESTART
; AS WHEN USE STARTMENU TO SHUTDOWN RESTART THAT SEEM TO STAY THERE 
; WHILE QUESTION ARE GIVEN ABOUT PROCESS CLOSING
; I FIND IT VERY ANNOYER
;# ------------------------------------------------------------------
; Thu 29-Nov-2018 14:59:43
; Thu 29-Nov-2018 16:04:00
;# ------------------------------------------------------------------

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
; https://www.dropbox.com/s/bsvrrmyknxr7pvd/Autokey%20--%2018-NORTON%20CONTROL%20BOOTER.ahk?dl=0
;# ------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

EXITAPP
Process, Close,% DllCall("GetCurrentProcessId")

; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp HAVE AR CALL ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------





SetStoreCapslockMode, off

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


OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

GLOBAL SIGNAL_TO_RESTART_HAPPEN
GLOBAL I_COUNT

SIGNAL_TO_RESTART_HAPPEN=FALSE
I_COUNT=0

GLOBAL ESCAPE_KEY_A_FEW
ESCAPE_KEY_A_FEW=0

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

DetectHiddenWindows, ON

; PAUSE

settimer MAIN_RUNNER, 1000

; settimer WINDOWS_10_STATRT_MENU_DOWN, 100
; settimer WINDOWS_10_STATRT_MENU_DOWN, off

RETURN  

WINDOWS_10_STATRT_MENU_DOWN:
	window=ahk_class Windows.UI.Core.CoreWindow
	isWindowShow(window)
RETURN    

;; F1::EXITAPP

;; F1::Process, Close,% DllCall("GetCurrentProcessId")

MAIN_RUNNER:

	SET_GO=TRUE
	If WinExist("SystemExplorer")
		SET_GO=FALSE
	If WinExist("CAsyncSocketEx Helper Window")
		SET_GO=FALSE
	If WinExist("End Program - CSR_SYNCML_CLASS_1EF5ED00AB77")
		SET_GO=FALSE

	IF SIGNAL_TO_RESTART_HAPPEN=TRUE
	{
		SoundBeep , 2000 , 100
		ESCAPE_KEY_A_FEW+=1
		IF ESCAPE_KEY_A_FEW<100
		{
			; CODE HELP CREDIT 
			; ----
			; Taskbar and Start Menu manipulation - AutoHotkey Community
			; https://www.autohotkey.com/boards/viewtopic.php?t=37718
			; ----
			IF (OSVER_N_VAR = 10) ; WIN 10
			{
				fVisible=0
				AppVisibility := ComObjCreate(CLSID_AppVisibility := "{7E5FE3D9-985F-4908-91F9-EE19F9FD1514}", IID_IAppVisibility := "{2246EA2D-CAEA-4444-A3C4-6DE827E44313}")
				if (DllCall(NumGet(NumGet(AppVisibility+0)+4*A_PtrSize), "Ptr", AppVisibility, "Int*", fVisible) >= 0)
				IF fVisible=1 
				{
					Send {LWin}
				}
			}
		}
		
		Process, Close, SystemExplorer.exe
		Process, Close, FileZilla Server Interface.exe
		; Process, Close, FileZilla Server Interface.exe
		
	}
	
	; ----------------------------------------------------
	; IF SystemExplorer.exe              DOESN'T EXIST OR
	; IF FileZilla Server Interface.exe  DOESN'T EXIST 
	; THEN NOT ANY POINT BEING HERE AFTER A WHILE
	; EXAMPLE LIKE BOOT UP
	; ----------------------------------------------------
	I_COUNT+=1
	IF I_COUNT<480
		SET_GO=FALSE


	IF SET_GO=TRUE
	{
		SoundBeep , 2000 , 100
		SoundBeep , 2500 , 100
		; KILL ITSELF
		Process, Close,% DllCall("GetCurrentProcessId")
	}

	;---------------------------------------------------
	; THESE PAIR WON'T REALLY GET RUN BUT LEFT IN ANYWAY
	; AS WORK TO FIND OUT FIRST OFF
	; PRIORITY IS IN THE EXITAPP
	;---------------------------------------------------

	;--------------------------------------
	;C:\PStart\Progs\#_PortableApps\PortableApps\SystemExplorerPortable\App\SystemExplorer\SystemExplorer.exe	
	;--------------------------------------
	IfWinExist End Program - SystemExplorer	
	{
		; Run, "TASKKILL.exe" /F /IM SystemExplorer.exe /T , , HIDE
		Sleep 8000
		SoundBeep , 2500 , 100
		ControlClick, &End Now, End Program - SystemExplorer
	}

	;------------------------------
	;FileZilla Server Interface.exe
	;CAsyncSocketEx Helper Window
	;------------------------------
	IfWinExist End Program - CAsyncSocketEx Helper Window
	{
		;WinGet, path, ProcessName, CAsyncSocketEx Helper Window
		Sleep 18000
		SoundBeep , 2500 , 100
		ControlClick, &End Now, End Program - CAsyncSocketEx Helper Window
	}	

	IfWinExist End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
	{
		; CODE HELP CREDIT 
		; ----
		; Taskbar and Start Menu manipulation - AutoHotkey Community
		; https://www.autohotkey.com/boards/viewtopic.php?t=37718
		; ----
		IF (OSVER_N_VAR = 10) ; WIN 10
		{
			fVisible=0
			AppVisibility := ComObjCreate(CLSID_AppVisibility := "{7E5FE3D9-985F-4908-91F9-EE19F9FD1514}", IID_IAppVisibility := "{2246EA2D-CAEA-4444-A3C4-6DE827E44313}")
			if (DllCall(NumGet(NumGet(AppVisibility+0)+4*A_PtrSize), "Ptr", AppVisibility, "Int*", fVisible) >= 0)
			IF fVisible=1 
			{
				Send {LWin}
			}
		}

		Sleep 4000
		SoundBeep , 2500 , 100
		WINACTIVATE, End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
		ControlClick, &End Now, End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
	}	

	
RETURN



; ------------------------------------------------------------------
isWindowShow( winTitle ) {
; checks if the specified window is full screen

DetectHiddenWindows, ON

winID := WinExist( winTitle )

If ( !winID )
	Return false

winW=
winH=
	
WinGet style, Style, ahk_id %WinID%
WinGetPos ,,x,y,winW,winH, ahk_id %WinID%

; 0x800000 is WS_BORDER.
; 0x20000000 is WS_MINIMIZE.
; no border and not minimized

tooltip % x " -- " y " -- "winW " -- " winH " -- " style


Return ((style & 0x20800000) 
or winH < A_ScreenHeight 
or winW < A_ScreenWidth) ? false : true

; ----
; Detect Fullscreen application? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/38882-detect-fullscreen-application/
; ----
}


;# ------------------------------------------------------------------





#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


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

    if ExitReason in Logoff,Shutdown
    {
		
		SIGNAL_TO_RESTART_HAPPEN=TRUE
	
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
		
		;---------------------------------------------------------------------
		; TASKKILL.exe WON'T WORK CAN'T LAUNCH ANOTHER PROGRAM DURING SHUTDOWN
		;---------------------------------------------------------------------
		;Run, "TASKKILL.exe" /F /IM SystemExplorer.exe /T , , MIN
		;Run, "TASKKILL.exe" /F /IM FileZilla Server Interface.exe /T , , MIN
		
		;---------------------------------------------------------------------
		; WaitClose DOSEN;T WORK HER REM LINE
		; MUST BE KILLED INSTANTLY
		;---------------------------------------------------------------------
		; Process, WaitClose, ahk_exe SystemExplorer.exe, 2
		; Process, WaitClose, ahk_exe FileZilla Server Interface.exe, 2
		;---------------------------------------------------------------------
		
		;---------------------------------------------------------------------
		; THE PROBLEM IS ABOUT REQUIRING THIS WAY
		; IS IF THIS PROGRAM GET REQUEST TO SHUTDOWN BEFORE THE OTHER FEW COUPLE
		; AND THEN A HOLD UP PROBLEM
		; SO HAVE TO BE KILLED AS SOON AS SHUTDOWN REQUEST COMES ALONG
		;---------------------------------------------------------------------
		
		SoundBeep , 2000 , 100
		SENDINPUT {ESC}

		
		Process, Close, SystemExplorer.exe
		Process, Close, FileZilla Server Interface.exe
		
		; Process, Close, C:\Program Files\Logitech\SetPointP\Campaign\LogiCampaignNotifier.exe
		Process, Close, LogiCampaignNotifier.exe
	
		SoundBeep , 2000 , 100
		SENDINPUT {ESC}
	
		;---------------------------------------------------------------------
		; THE RETURN 1 IS STILL USED BUT THE APP CLOSES ITSELF 
		; PROBABLY BECAUSE THE COUPLE OF PROGRAM ARE FOUND NOT TO EXIST
		; BUT EVEN WITH A MSGBOX IN THE WAY IT STILL GET CLOSE
		;---------------------------------------------------------------------
		
		; IF EXIT_APP_VAR=FALSE 
			; {
			; ; MSGBOX % EXIT_APP_VAR
			; return 1
			; }
	}
	
	; PROCESS SCRIPT HERE WILL ONLY CLOSE BY FORCE KILL _ DONE WITHIN
	RETURN 1
	
	; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
		; THIS ROUTINE WON'T GET CALLED UNLESS ExitFunc HAS RETURN CLEAR TO CLOSE PREVENT WITH RETURN 1
		SoundBeep , 2500 , 100
		SoundBeep , 3000 , 100
		
		; MsgBox, MyObject is cleaning up prior to exiting...

		;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}

; -----------------------------------------------------------------
; exit the app
