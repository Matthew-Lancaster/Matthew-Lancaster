; ----------------------------------------------------------------
; __ MATTHEW LANCASTER 
; __ MATT.LAN@BTINTERNET.COM
; ----------------------------------------------------------------
; __ FROM ---- Sun 27-May-2018 13:00:00
; __ TO   ---- Sun 27-May-2018 13:02:00 __ 2 MINUTE
; ----------------------------------------------------------------

;-----------------------------------------------------------------
; HIKVISION CLOUD GOOGLE DRIVE IMAGES SNAPSHOT 
; CONVERT WITH IrfanView TO MAKE GOOGLE ABLE TO DISPLAY
;-----------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance ignore

SET_GO=TRUE
Process, Exist, wscript.exe
If ErrorLevel > 0 ; errorlevel will = 0 if process doesn't exist
	{
	SET_GO=FALSE 
	}
Process, Exist, i_view32.exe
If  ErrorLevel > 0 ; errorlevel will = 0 if process doesn't exist
{	
	SET_GO=FALSE 
}

if ScriptInstanceExist()
{
	Exitapp
}

IF (A_ComputerName = 7-ASUS-GL522VW) 
{
	Exitapp
}


;--------------------------------------------------------------------
IF SET_GO=TRUE 
{
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 24-i_view32 convert_CCSE.VBS"
IfExist, %FN_VAR%
	RunWait, "%FN_VAR%" , , HIDE
}


Exitapp


ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return
