;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 34-CPC CHROME PAGE 404 REMOVE.ahk
;# __ 
;# __ Autokey -- 34-CPC CHROME PAGE 404 REMOVE.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# START TIME [ Fri 29-Jun-2018 06:29:35 ]
;# END   TIME [ Fri 29-Jun-2018 07:19:41 ]
;# __ 
;====================================================================

;# ------------------------------------------------------------------
; DESCRIPTION 
;# ------------------------------------------------------------------
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
; 
;# ------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100
SetStoreCapslockMode, off

DetectHiddenWindows, ON

WinGet, list, List, ahk_exe Chrome.exe
LOOP {
	RE_DO_LOOP=FALSE
	Loop %list% {
		hwnd := list%A_Index%
		SetTitleMatchMode, 2
		WinGetTitle, WLName, ahk_id %hwnd%
		SET_GO=FALSE
		if(InStr(WLName, "404 Page Not Found | CPC")>0)
			SET_GO=TRUE
		XX_VAR:="New Tab - Google Chrome"
		if WLName=%XX_VAR%
		{
			SET_GO=TRUE
			msgbox hh
		}
		if WLName=CPC - Google Chrome
		{
			SET_GO=TRUE
		}
		
		IF SET_GO=TRUE 
		{
			WinActivate, ahk_id %hwnd%
			SENDINPUT, ^w
			RE_DO_LOOP=TRUE
			sleep 400
		}
		sleep 20
	}

	IF RE_DO_LOOP=FALSE
	BREAK
} 




;------------------------------------------------------------------
; CREDIT
; BUT REQUIRE VERY MUCH IMPROVEMENT
; SEEN THIS EXAMPLE BEFORE AND ALL OF THEM REQUIRE IT IMPROVER
;------------------------------------------------------------------
; ----
; Cycle through chrome windows and then tabs - AutoHotkey Community
; https://autohotkey.com/boards/viewtopic.php?t=17674
; ----
;------------------------------------------------------------------
