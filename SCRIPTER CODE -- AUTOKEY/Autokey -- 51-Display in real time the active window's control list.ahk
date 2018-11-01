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
;IT USER ExitFunc TO EXIT FROM #Persistent
;OR      Exitapp  TO EXIT FROM #Persistent
;Exitapp CALLS ONTO ExitFunc
;--------------------
;--------------------
#SingleInstance force
;--------------------

SetTimer, WatchActiveWindow_2, 200
return

WatchActiveWindow_1:
WinGet, ControlList, ControlList, A
ToolTip, %ControlList%
return

WatchActiveWindow_2:
WinGet, ControlList, ControlList, ahk_class CabinetWClass
ToolTip, %ControlList%

; NEXT WANTED TITLE FROM EACH CONTROL ALONGSIDE EXTENDED CLASS NAME

return