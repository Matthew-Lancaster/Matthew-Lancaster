;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 51-Display in real time the active window's control list.ahk
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
; ---- Location Internet
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01-INCLUDE MENU 01 of 03.ahk


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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02-INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03-INCLUDE MENU 03 of 03.ahk
