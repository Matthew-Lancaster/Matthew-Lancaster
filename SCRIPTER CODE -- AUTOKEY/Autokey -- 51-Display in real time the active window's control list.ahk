;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 51-Display in real time the active window's control list.ahk
;# __ 
;# __ Autokey -- 51-Display in real time the active window's control list.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ 
;# __ FROM ---- Thu 01-Nov-2018 09:24:49 __ CODE KNOCK UP TIME AND WORKED ON OTHER THING FOR HOUR NESTING WORKER
;# __ TO   ---- Thu 01-Nov-2018 11:08:00
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; ---- LOCATION ON-LINE
; -------------------------------------------------------------------
; ---- WWW.GITHUB.COM
;
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; SOURCE MY SEARCHER
; ----
; AHK LIST ALL CHILD WINDOW site:autohotkey.com - Google Search
; https://www.google.co.uk/search?q=AHK+LIST+ALL+CHILD+WINDOW+site:autohotkey.com&num=50&safe=active&sa=X&ved=2ahUKEwi7id3D77LeAhVLBsAKHRYuAD4QrQIoBDAAegQIAxAN&biw=1536&bih=551
; ----
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SOURCE CREDIT 
; ----
; WinGet - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/WinGet.htm
; ----
; -------------------------------------------------------------------


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

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

;--------------------
#SingleInstance force
;--------------------

SetTitleMatchMode, 3
; SetTitleMatchMode, slow

setTimer, WatchActiveWindow_2, 200

; SetTimer, WatchCursor, 50

return

; ESC::EXITAPP
ESC::
{
tooltip
pause
}

WatchActiveWindow_1:
WinGet, ControlList, ControlList, A
ToolTip, %ControlList%
return

WatchActiveWindow_2:

; WinGet, ControlList, ControlList, ahk_class CabinetWClass
; WinGet, ControlList, ControlList, A
WinGet, ControlList, ControlList, ahk_class Chrome_WidgetWin_1

; WinGet, ControlList, ControlList, ahk_class ThunderRT6FormDC
; WinGet, ControlList, ControlList, "EliteSpy+ by Andrea B 2001 __ www.PlanetSourceCode.com_ & Big Timer Worker By Matthew Lancaster __ 07722224555 __ Version 1.0.421"

Loop, Parse, ControlList, `n

{

	ClassNN := A_LoopField

	; ControlGetPos, X, Y, W, H, %ClassNN%, ahk_class CabinetWClass
	; ControlGetText, OutputVar, %ClassNN%, ahk_class CabinetWClass
	
	ControlGetText, OutputVar, %ClassNN%, ahk_class Chrome_WidgetWin_1

	; Controls .= ClassNN "`t" X "," Y " - " W "," H "`n"
	Controls .= ClassNN "`t" OutputVar "`n"
	; Controls .= ClassNN "`n"

}

ToolTip, %ControlS%
; ToolTip, %ControlList%

clipboard = %ControlS%
setTimer, WatchActiveWindow_2, off

; NEXT WANTED TITLE FROM EACH CONTROL ALONGSIDE EXTENDED CLASS NAME

return


WatchCursor:
MouseGetPos, X, Y, WindowId, ControlId, 

ControlGetText, Text, %ControlId%

X += 10
Y += 10

ToolTip, Window Title: %WindowId% `n Text under mouse: ^%Text%^%ControlId%, %X%, %Y%
return



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
