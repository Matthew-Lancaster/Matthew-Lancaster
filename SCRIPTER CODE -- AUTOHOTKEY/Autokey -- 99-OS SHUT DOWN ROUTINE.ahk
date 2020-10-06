
; The following DllCall is optional: it tells the OS to shut down this script first (prior to all other applications).
DllCall("kernel32.dll\SetProcessShutdownParameters", "UInt", 0x4FF, "UInt", 0)
OnMessage(0x11, "WM_QUERYENDSESSION")



 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
;# __ 
;# __ Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
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
#SingleInstance force
; -------------------------------------------------------------------

SET_SHUT_DOWN=

; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SETSTORECAPSLOCKMODE, OFF


RETURN


#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


; -------------------------------------------------------------------
; OnMessage() - Syntax & Usage | AutoHotkey 
; https://www.autohotkey.com/docs/commands/OnMessage.htm#shutdown
; ----
; Sun 04-Oct-2020 12:34:26
; -------------------------------------------------------------------

WM_QUERYENDSESSION(wParam, lParam)
{
    ENDSESSION_LOGOFF := 0x80000000
    if (lParam & ENDSESSION_LOGOFF)  ; User is logging off.
	{
        EventType := "Logoff"
		SET_SHUT_DOWN=TRUE
	}
    else  ; System is either shutting down or restarting.
	{
        EventType := "Shutdown"
		SET_SHUT_DOWN=TRUE
	}

	IF SET_SHUT_DOWN=TRUE
	{
		GOSUB TERMINATE_ALL_AUTOHOTKEYS_GONE_NOT_INCLUDER
		GOSUB KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT_NOT_INCLUDER
		GOSUB TERMINATE_ALL_AUTOHOTKEYS_GONE_NOT_INCLUDER
		
		; ---------------------------------------------------------------
		; LET OWN ONE EXIT STYLE -- ICON GONE
		; ---------------------------------------------------------------
		SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")
		PROCESS, Close, %SCRIPTOR_OWN_PID% 
		; ---------------------------------------------------------------

		EXITAPP
		RETURN
	}
	

    ; try
    ; {
        ; ; Set a prompt for the OS shutdown UI to display.  We do not display
        ; ; our own confirmation prompt because we have only 5 seconds before
        ; ; the OS displays the shutdown UI anyway.  Also, a program without
        ; ; a visible window cannot block shutdown without providing a reason.
        ; BlockShutdown("Example script attempting to prevent " EventType ".")
        ; return false
    ; }
    ; catch
    ; {
        ; ; ShutdownBlockReasonCreate is not available, so this is probably
        ; ; Windows XP, 2003 or 2000, where we can actually prevent shutdown.
        ; MsgBox, 4,, %EventType% in progress.  Allow it?
        ; IfMsgBox Yes
            ; return true  ; Tell the OS to allow the shutdown/logoff to continue.
        ; else
            ; return false  ; Tell the OS to abort the shutdown/logoff.
    ; }
}

BlockShutdown(Reason)
{
    ; If your script has a visible GUI, use it instead of A_ScriptHwnd.
    DllCall("ShutdownBlockReasonCreate", "ptr", A_ScriptHwnd, "wstr", Reason)
    OnExit("StopBlockingShutdown")
}

StopBlockingShutdown()
{
    OnExit(A_ThisFunc, 0)
    DllCall("ShutdownBlockReasonDestroy", "ptr", A_ScriptHwnd)
}




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
				PROCESS, Close, %PID_8% 
				SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
				; SOUNDBEEP 1000,20
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
	; ----------------------------------------------------
	FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk"
	IfExist, %FN_VAR_1%
	{
		Run, %FN_VAR_1%
		WINWAIT Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk ahk_class AutoHotkey,,30
	}

	; -------------------------------------------------------------------
	; KILL ALL VISUAL BASIC AND COMMAND
	; -------------------------------------------------------------------
	; GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK
	GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_GONE
	; ---------------------------------------------------------------
	; GOSUB KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT

	; ---------------------------------------------------------------
	; LET OWN ONE EXIT STYLE -- ICON GONE
	; ---------------------------------------------------------------
	; SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")
	; PROCESS, Close, %SCRIPTOR_OWN_PID% 
	; ---------------------------------------------------------------
	
	EXITAPP
	RETURN
		
RETURN



KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT_NOT_INCLUDER:
	WinGet, List, List, ahk_exe CMD.EXE
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
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
			Process, Close, %PID_8% 
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
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	
RETURN



CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_GONE_NOT_INCLUDER:
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
							Process, Close, %PATH_EXE%
						}
					}
				}

			}
	}	
	TOOLTIP
	SLEEP 500
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
RETURN


