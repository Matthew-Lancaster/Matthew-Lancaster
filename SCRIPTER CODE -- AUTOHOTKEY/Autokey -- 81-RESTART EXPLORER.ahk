;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 81-RESTART EXPLORER.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; ONLY ONE THING 
; AND MSGBOX WITH OPTION COUPLE ARE SUM OF EACH OTHER
; FIND HARD PUT IN VARIABLE AND INCLUDE THAT WAY
; GAVE UP IN END - CALCULATOR AT THE READY
;
; -------------------------------------------------------------------
; FROM __ Tue 10-Sep-2019 14:42:00
; TO   __ Tue 10-Sep-2019 15:42:00
; -------------------------------------------------------------------

#noEnv
; #persistent
#singleInstance force
detectHiddenWindows, on
setWorkingDir %a_scriptDir%
; #NoTrayIcon

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

GOSUB GO_ROUTINE

EXITAPP

GO_ROUTINE:

	; WIN_XP 5 WIN_7 6 WIN_10 10  
	; --------------------------
	OSVER_N_VAR:=a_osversion
	IF INSTR(a_osversion,".")>0
		OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
	IF OSVER_N_VAR=WIN_XP
		OSVER_N_VAR=5
	IF OSVER_N_VAR=WIN_7
		OSVER_N_VAR=6

	IF OSVER_N_VAR<10
		EXITAPP
		
	; ---------------------------------------------------------------
	; MAKE HERE ROUTINE AS VBASIC WON'T DISPLAY 
	; THAT EASIER A COUNTDOWN MSGBOX
	; ---------------------------------------------------------------
	
	SoundBeep , 1500 , 400
	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Process, Exist, explorer.exe
	If ErrorLevel > 0
	{
		M2=4
		M2+=4096
		M2=4100    ; DISPLAY ALWAYS ON TOP AND ASK QUESTION YES NOT
		FormatTime, TimeString, HH:mm, HH:mm
		FormatTime, TimeString, HH, HH
		IF TimeString=00
		LINE_STRING=RESTART EXPLORER FOR MIDNIGHT`n`n10 SECOND TO ANSWER NOT`n`nDEFAULT YES
		IF TimeString<>00
		LINE_STRING=RESTART EXPLORER`n`n10 SECOND TO ANSWER NOT`n`nDEFAULT YES
		
		MSGBOX ,4100,,%LINE_STRING%,10
		IFMSGBOX NO
			EXITAPP

		Process,close,explorer.exe
	}
	
	; ---------------------------------------------------------------
	; REQUEST KILL EXPLORER IF NONE EXIST ABOVE 
	; AND THEN START IT GOING AGAIN
	; ---------------------------------------------------------------
	RUN, explorer.exe
	
	WinWait, ahk_class CabinetWClass
	WinClose ;close the new explorer window
		
	EXITAPP
	

RETURN


; -------------------------------------------------------------------
; REFERENCE FIND CREDIT
; -------------------------------------------------------------------
; ----
; AHK BEST WAY TO KILL EXPLORER - Google Search
; https://www.google.com/search?q=AHK+BEST+WAY+TO+KILL+EXPLORER
; --------
; Restart Explorer (the official way) - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/93906-restart-explorer-the-official-way/
; ----
