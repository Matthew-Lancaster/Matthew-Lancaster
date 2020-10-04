
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

EventType=

; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SetStoreCapslockMode, off


return


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
	}
    else  ; System is either shutting down or restarting.
	{
        EventType := "Shutdown"
	}

	IF EventType
	{
		GOSUB TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME
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