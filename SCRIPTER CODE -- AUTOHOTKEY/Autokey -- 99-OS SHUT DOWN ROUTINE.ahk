;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 99-OS SHUT DOWN ROUTINE.ahk
;# __ 
;# __ Autokey -- 99-OS SHUT DOWN ROUTINE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; WORK TIME 2ND DAY
; -------------------------------------------------------------------
; Tue 06-Oct-2020 14:18:17
; Tue 06-Oct-2020 16:48:00 -- 02 HOUR 29 MINUTE
; -------------------------------------------------------------------




; The following DllCall is optional: it tells the OS to shut down this script first (prior to all other applications).
DllCall("kernel32.dll\SetProcessShutdownParameters", "UInt", 0x4FF, "UInt", 0)
OnMessage(0x11, "WM_QUERYENDSESSION")

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; -------------------------------------------------------------------
; #SINGLEINSTANCE FORCE
; -------------------------------------------------------------------
#PERSISTENT
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

SETSTORECAPSLOCKMODE, OFF

; -------------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 0001 SOMETHING DO LOGGOFF SHUTDOWN
; 0002 APP IT OWN GET REQUEST CLOSE WITH BOOT LOGGOFF SHUTDOWN
; -------------------------------------------------------------------


SET_SHUT_DOWN=


; -------------------------------------------------------------------
; END CODE BLOCK INIT
; -------------------------------------------------------------------
RETURN
; -------------------------------------------------------------------



#INCLUDE C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk



; -------------------------------------------------------------------
; -------------------------------------------------------------------
SHUTDOWN_ROUTINE:
	PROCESS, CLOSE, SYSTEMEXPLORER.EXE
	PROCESS, CLOSE, FILEZILLA SERVER INTERFACE.EXE
	; C:\PROGRAM FILES\LOGITECH\SETPOINTP\CAMPAIGN\
	PROCESS, CLOSE, LOGICAMPAIGNNOTIFIER.EXE

	; ---------------------------------------------------------------
	; TWO MINUTE
	; 4 SECOND WHEN ALL CLEAR
	; BY SETTIMER MAIN_RUNNER, 400
	; ---------------------------------------------------------------
	
	
	
	SETTIMER TIMER_MSGBOX_KEEP_WAIT_OWN_SCRIPT_POP_AWAY,400
	SETTIMER TIMER_SET_SHUTDOWN_DO,1200000    

	SETTIMER WINDOWS_10_STATRT_MENU_DOWN      , 1000
	SETTIMER WINDOWS_STATRT_MENU_DOWN_GENERAL , 1000
	SETTIMER MAIN_RUNNER, 400
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	
	GOSUB TERMINATE_ALL_AUTOHOTKEYS_GONE_NOT_INCLUDER
	GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_GONE_NOT_INCLUDER
	GOSUB KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT_NOT_INCLUDER

RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
TIMER_MSGBOX_KEEP_WAIT_OWN_SCRIPT_POP_AWAY:

	WinGet, VAR_GET, ID, Autokey -- 99-OS SHUT DOWN ROUTINE.ahk ahk_class #32770
	IF !VAR_GET
		RETURN

	; ---------------------------------------------------------------
	; HERE RELOAD WITH MY NEW ROUTINE
	; RATHER THAN OLD ONE LEFT THERE
	; ---------------------------------------------------------------
	WinGet, VAR_GET, PID, Autokey -- 99-OS SHUT DOWN ROUTINE.ahk ahk_class #32770
	PROCESS, CLOSE, %VAR_GET% 
	RETURN

	; ---------------------------------------------------------------
	; RATHER THAN OLD ONE LEFT THERE
	; ---------------------------------------------------------------
	ControlClick, Button2, ahk_id %VAR_GET%,,,, NA x10 y10
	ControlClick, Button2, ahk_id %VAR_GET%

	; ---------------------------------------------------------------
	; Text:	&No
	; ClassNN:	Static1
	; Text:	Could not close the previous instance of (...)
	; ---------------------------------------------------------------
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
TIMER_SET_SHUTDOWN_DO:
	; KILLER ITSELF
	PROCESS, CLOSE,% DllCall("GetCurrentProcessId")
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
MAIN_RUNNER:
	
	; ---------------------------------------------------------------
	ALL_CLEAR_SHUTDOWN=TRUE
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	If WinExist("SystemExplorer")
		ALL_CLEAR_SHUTDOWN=FALSE
	If WinExist("CAsyncSocketEx Helper Window")
		ALL_CLEAR_SHUTDOWN=FALSE
	If WinExist("End Program - CSR_SYNCML_CLASS_1EF5ED00AB77")
		ALL_CLEAR_SHUTDOWN=FALSE
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	; KILLER IT OWN
	; ---------------------------------------------------------------
	IF ALL_CLEAR_SHUTDOWN=TRUE
	{
		SETTIMER TIMER_SET_SHUTDOWN_DO,4000
		SETTIMER MAIN_RUNNER,OFF
	}
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	; THESE PAIR WON'T REALLY GET RUN BUT LEFT IN ANYWAY
	; AS WORK TO FIND OUT FIRST OFF
	; PRIORITY IS IN THE EXITAPP
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	;C:\PStart\Progs\#_PortableApps\PortableApps\SystemExplorerPortable\App\SystemExplorer\SystemExplorer.exe	
	; ---------------------------------------------------------------
	IfWinExist End Program - SystemExplorer	
	{
		; Run, "TASKKILL.exe" /F /IM SystemExplorer.exe /T , , HIDE
		Sleep 8000
		SoundBeep , 2500 , 100
		ControlClick, &End Now, End Program - SystemExplorer
	}
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	;FileZilla Server Interface.exe
	;CAsyncSocketEx Helper Window
	; ---------------------------------------------------------------
	IfWinExist End Program - CAsyncSocketEx Helper Window
	{
		;WinGet, path, ProcessName, CAsyncSocketEx Helper Window
		Sleep 18000
		SoundBeep , 2500 , 100
		ControlClick, &End Now, End Program - CAsyncSocketEx Helper Window
	}	
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
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
		SLEEP 4000
		SOUNDBEEP , 2500 , 100
		WINACTIVATE, END PROGRAM - CSR_SYNCML_CLASS_1EF5ED00AB77
		CONTROLCLICK, &END NOW, END PROGRAM - CSR_SYNCML_CLASS_1EF5ED00AB77
	}	
	; ---------------------------------------------------------------
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
TERMINATE_ALL_AUTOHOTKEYS_GONE_NOT_INCLUDER:
	DETECTHIDDENWINDOWS, ON
	SETTITLEMATCHMODE, 2
	
	SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")

	; -------------------------------------------------------------------
	; KILL ALL NOT INCLUDE OWN OR TRAY_ICON_CLEANER
	; -------------------------------------------------------------------
	WinGet, List, List, ahk_class AutoHotkey
	PID_8=
	Loop
	{
		Loop %List%
		{
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			IF PID_8
			If PID_8 <> %SCRIPTOR_OWN_PID%
			IF PID_8 <> %PID_78_TRAY_ICON_CLEANER%
			{
				PROCESS, CLOSE, %PID_8% 
				SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
				; SOUNDBEEP 1000,20
				SOME_AHK_PROCCESS_CLOSE=TRUE
			}
		}
		WinGet, List, List, ahk_class AutoHotkey
		Loop %List%
			I_COUNT=%A_Index%
		SLEEP 50
		IF I_COUNT<3 THEN 
			BREAK
	}

	; RUN HERE AS ANOTHER TYPE SCRIPT PROGRAMMER LANGUAGE
	; TRAY CLEANER
	; ----------------------------------------------------
	FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk"
	IF SOME_AHK_PROCCESS_CLOSE=TRUE
	IfExist, %FN_VAR_1%
	{
		Run, %FN_VAR_1%
		WINWAIT Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk ahk_class AutoHotkey,,30
	}
	; ---------------------------------------------------------------
	; LET OWN ONE EXIT STYLE -- ICON GONE
	; ---------------------------------------------------------------
	; SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")
	; PROCESS, CLOSE, %SCRIPTOR_OWN_PID% 
	; ---------------------------------------------------------------

RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT_NOT_INCLUDER:
	WinGet, List, List, ahk_exe CMD.EXE
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			PROCESS, CLOSE, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	WinGet, List, List, ahk_exe CONHOST.EXE
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			PROCESS, CLOSE, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	WinGet, List, List, ahk_exe WSCRIPT.EXE	
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			PROCESS, CLOSE, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_GONE_NOT_INCLUDER:
	; ---------------------------------------------------------------
	DetectHiddenWindows, ON
	WinGet, List, List, ahk_class ThunderRT6FormDC
	PATH_ID_BUILD=
	Loop %List%  
	{
		; IfWinExist ahk_id List%A_Index%
		;WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		WinGet, PATH_FULL, ProcessPath, % "ahk_id " List%A_Index% 
		WinGet, PATH_EXE, ProcessName, % "ahk_id " List%A_Index% 
		StringUpper PATH_EXE, PATH_EXE
		StringUpper PATH_FULL, PATH_FULL
			IF INSTR(PATH_FULL,"D:\VB")>0
			{
			
				IF INSTR(PATH_ID_BUILD,PATH_EXE)=0
				{
					PATH_ID_BUILD=%PATH_ID_BUILD%%PATH_EXE%`n
					PROCESS, EXIST, %PATH_EXE%
					IF ERRORLEVEL
					{
						TOOLTIP % PATH_ID_BUILD
						SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
						WINCLOSE ahk_exe %PATH_EXE%
						SLEEP 200
						PROCESS, EXIST, %PATH_EXE%
						IF ERRORLEVEL
						{
							SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							WINCLOSE ahk_exe %PATH_EXE%
						}
						SLEEP 200
						PROCESS, EXIST, %PATH_EXE%
						IF ERRORLEVEL
						{
							SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							PROCESS, CLOSE, %PATH_EXE%
						}
					}
				}

			}
	}	
	TOOLTIP
	SLEEP 500
	; ---------------------------------------------------------------
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
WINDOWS_10_STATRT_MENU_DOWN:
	; ---------------------------------------------------------------
	SETTIMER WINDOWS_10_STATRT_MENU_DOWN, 4000
	; ---------------------------------------------------------------
	window=ahk_class Windows.UI.Core.CoreWindow
	isWindowShow(window)
	; ---------------------------------------------------------------
	SENDINPUT {ESC}
	; ---------------------------------------------------------------
RETURN    
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
WINDOWS_STATRT_MENU_DOWN_GENERAL:
	; ---------------------------------------------------------------
	SETTIMER WINDOWS_STATRT_MENU_DOWN_GENERAL , 4000
	; ---------------------------------------------------------------
	; CODE HELP CREDIT 
	; ---------------------------------------------------------------
	; Taskbar and Start Menu manipulation - AutoHotkey Community
	; https://www.autohotkey.com/boards/viewtopic.php?t=37718
	; ---------------------------------------------------------------
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
	; ---------------------------------------------------------------
RETURN    
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
EOF:                           ; ON EXIT
EXITAPP     
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; BEGIN FUNCTION BLOCK
; -------------------------------------------------------------------
; -------------------------------------------------------------------
EXITFUNC(EXITREASON, EXITCODE)
{
    ; IF EXITREASON NOT IN LOGOFF,SHUTDOWN
		; SET_SHUT_DOWN=NOT A SHUTDOWN METHOD
		
    IF EXITREASON IN LOGOFF,SHUTDOWN
		SET_SHUT_DOWN=TRUE
	
	IF SET_SHUT_DOWN=TRUE
		GOSUB SHUTDOWN_ROUTINE
	
	; PROCESS SCRIPT HERE WILL ONLY CLOSE BY FORCE KILL _ DONE WITHIN
	; ONEXIT FUNCTIONS MUST RETURN NON-ZERO TO PREVENT EXIT.
	RETURN 1
	
	; DO NOT CALL EXITAPP -- THAT WOULD PREVENT OTHER ONEXIT FUNCTIONS FROM BEING CALLED.
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
class MyObject
{
    Exiting()
    {
		; THIS ROUTINE WON'T GET CALLED UNLESS EXITFUNC HAS RETURN CLEAR TO CLOSE PREVENT WITH RETURN 1
		; ---------------------------------------------------------------------------------------------
		; MSGBOX, MYOBJECT IS CLEANING UP PRIOR TO EXITING...
		; ---------------------------------------------------------------------------------------------
        /*
			THIS.SAYGOODBYE()
			THIS.CLOSENETWORKCONNECTIONS()
        */
    }
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; OnMessage() - Syntax & Usage | AutoHotkey 
; https://www.autohotkey.com/docs/commands/OnMessage.htm#shutdown
; ----
; Sun 04-Oct-2020 12:34:26
; -------------------------------------------------------------------
WM_QUERYENDSESSION(wParam, lParam)
{
    ENDSESSION_LOGOFF := 0x80000000
    IF (lParam & ENDSESSION_LOGOFF)  ; USER IS LOGGING OFF.
	{
        EventType := "Logoff"
		SET_SHUT_DOWN=TRUE
	}
    ELSE  ; SYSTEM IS EITHER SHUTTING DOWN OR RESTARTING.
	{
        EventType := "Shutdown"
		SET_SHUT_DOWN=TRUE
	}

	IF SET_SHUT_DOWN=TRUE
	{
		GOSUB SHUTDOWN_ROUTINE
	}
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
BlockShutdown(Reason)
{
    ; If your script has a visible GUI, use it instead of A_ScriptHwnd.
    DllCall("ShutdownBlockReasonCreate", "ptr", A_ScriptHwnd, "wstr", Reason)
    OnExit("StopBlockingShutdown")
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
StopBlockingShutdown()
{
    OnExit(A_ThisFunc, 0)
    DllCall("ShutdownBlockReasonDestroy", "ptr", A_ScriptHwnd)
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
ISWINDOWSHOW( winTitle ) {
	; CHECK WINDOW FULL SCREEN

	DETECTHIDDENWINDOWS, ON

	WINID := WinExist( winTitle )

	IF ( !WINID )
		RETURN FALSE

	WINW=
	WINH=
		
	WINGET STYLE, STYLE, ahk_id %WinID%
	WINGETPOS ,,X,Y,WINW,WINH, AHK_ID %WINID%

	; 0X800000 IS WS_BORDER.
	; 0X20000000 IS WS_MINIMIZE.
	; NO BORDER AND NOT MINIMIZED

	; TOOLTIP % X " -- " Y " -- "WINW " -- " WINH " -- " STYLE

	RETURN ((STYLE & 0X20800000) 
	OR WINH < A_SCREENHEIGHT 
	OR WINW < A_SCREENWIDTH) ? FALSE : TRUE

	; ---------------------------------------------------------------
	; DETECT FULLSCREEN APPLICATION? - ASK FOR HELP - AUTOHOTKEY COMMUNITY
	; https://autohotkey.com/board/topic/38882-detect-fullscreen-application/
	; ---------------------------------------------------------------
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; END FUNCTION BLOCK
; -------------------------------------------------------------------

