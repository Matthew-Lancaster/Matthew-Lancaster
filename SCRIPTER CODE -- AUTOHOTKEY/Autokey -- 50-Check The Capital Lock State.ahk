﻿;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 50-Check The Capital Lock State.ahk
;# __ 
;# __ Autokey -- 50-Check The Capital Lock State.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ FROM Tue 30-Oct-2018 22:07:50 __ CODE KNOCK UP TIME
;# __ TO   Wed 31-Oct-2018 01:25:00
;# __ TO   Wed 31-Oct-2018 02:41:00 __ EXTENDED PLAYER 
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; ---- Location Internet
; -------------------------------------------------------------------
; ---- GITHUB.COM
; ----
; Matthew-Lancaster/Autokey -- 50-Check The Capital Lock State.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2050-Check%20The%20Capital%20Lock%20State.ahk
; ----
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; This Code Is Duplicated Coder Located at 
; Because I Already Have Enough AutoHotKeys Running at Moment
; Nice to Keep it Separate
; Well I Decide is Good Enough to Use as Own Desperate Code Program
; And use ignore mouse on keyboard hook
; -------------------------------------------------------------------
; Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
; -------------------------------------------------------------------

; # ------------------------------------------------------------------
; DESCRIPTION
; # ------------------------------------------------------------------
; AN INDICATOR FOR CAPS LOCK IS ON
; MY METHOD INTRO LEACHED CREDIT GIVEN SOURCE FROM
; HAS MY EXTRA SKILLFUL THOUGHT PUT IN
; THIS WILL DISPLAY A PROGRESS BAR BOTTOM RIGHT CORNER WITH INFO CAPS LOCK IF IT IS ON
; IT WILL SOUNDER AT CHANGE OF CAPS LOCK EITHER HI TON FOR UP AND LOW DOWN
; ALSO A SPECIAL BIT THAT OFTEN UPSET PEOPLE 
; WHEN RETURN TO KEYBOARD AFTER A MINUTE OR 20 SECONDS AWAY AND BEGIN KEY-ER
; IT WILL NOTIFY WITH SOUND BEEPER THAT ALSO HI OR LOW TONE
; I LIKED IT THAT WITH LOGITECH MOUSE 
; BUT THE SOFTWARE WORKED AT FIRST INSTALL AND NEVER AGAIN 
; THE HI LOW SOUND BUT MADE A ANNOYING BEEP INSTEAD
; IT ANNOYING ALSO WITH LOGITECH THE MOUSE DISPLAY OF AN ON-SCREEN BOX AT CAPS CHANGE
; CAN'T TURN IT OFF ANYMORE
; AND DOUBLE ANNOYING THE LOGITECH OPTIONS NEW VERSION OF SOFTWARE HAS ONE AS WELL
; MAKING TWO ALMOST AS IN CONFLICT
; WHEN YOU HAVE BOTH TYPE OF MOUSE AND THEIR SOFTWARE RUNNER
; -------------------------------------------------------------------
; Could Be Added to Press Shift on Own to Turn Caps Lock Off
; Test it Out Later
; -------------------------------------------------------------------
; Add Extra Shift Key Soundbeeper Only Sound if Caps Lock is On
; -------------------------------------------------------------------
; # ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; FROM Tue 30-Oct-2018 22:07:50 __ CODE KNOCK UP TIME
; TO   Wed 31-Oct-2018 01:25:00
; TO   Wed 31-Oct-2018 02:41:00 __ EXTENDED PLAYER 
; # ------------------------------------------------------------------


; # ------------------------------------------------------------------
; SESSION 002
; Fixed Finally Not to USE * in *Shift HotKey Make Repeat Soundbeep with Hold Shift Another Key if Faulty
; Removed Mouse Idle Trigger for Event Don't Want Reminded Caps Lock when Move Mouse
; After First Draft Copy  I Noticed Fault if Wasn't Corrector Made the control shift N Key Bring New Folder 
; in Explorer Go Wrong and Open New Explorer Up - Fixed Now
; -------------------------------------------------------------------
; FROM Wed 31-Oct-2018 12:45:22
; TO   Wed 31-Oct-2018 14:10:00
; # ------------------------------------------------------------------

; # ------------------------------------------------------------------
; SESSION 003
; WORK SOME COSMETIC IMPROVER NEW IDEA FOR
; AFTER CAPS LOCK PRESS WHICH MAKE SOUND AND THEN ANOTHER KEY PRESS WHICH ALSO MAKE SOUND ONCE
; TOOK LONG TIME UNTIL FIND GETLASTKEY A_PriorKey
; ALSO A CHANGE IN SHIFT KEY PRESS WITH LESS SOUND AND SIMILAR TO ABOVE WHERE ONLY ONE SOUND IS MADE AFTER CAPS LOCK KEY-PRESSER
; BUT FINDING HAVING SHIFT MAKE A SOUND IF A BIT MISLEADING BECAUSE PRESS SHIFT WILL MAKE OPPOSITE ACTION OF CAPS FOR FOLLOWING KEY-PRESS
; MUST WORK ON THIS 
; LIKE NEXT KEY PRESS MAKE A SOUND WITH TRUE REFLECTION 
; WOULD LIKE NEXT KEY PRESS DONE 1ST AND LOOK AT LATER REFLECTION IDEA
; -------------------------------------------------------------------
; FROM Tue 13-Nov-2018 22:27:07
; TO   WED 14-Nov-2018 01:28:00
; # ------------------------------------------------------------------

; # ------------------------------------------------------------------
; SESSION 004
; USE FULL SCREEN DETECTION BETTER _ DON'T DISPLAYER IN FULL SCREEN
; AND USE OS VERSION NUMBER _ TUINED UP FOR __ WIN XP WIN 07 WIN 10
; AND IN PROCESS CLEARED UP FAULT WITH WINDOWS XP DETECT DESKTOP 
; AS NAME PROGRAM MANAGER _ HELPER WITH FAULT IN THAT WHERE LEACHED FROM
; Autokey -- 14-BRIGHTNESS WITH DIMMER.ahk
; -------------------------------------------------------------------
; FROM Tue 26-Feb-2019 12:13:29
; TO   Tue 26-Feb-2019 12:13:29
; # ------------------------------------------------------------------



#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

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


; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------

; ----
; Another On-Screen Caps Lock Indicator - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/100990-another-on-screen-caps-lock-indicator/
; ----

#InstallKeybdHook  ;A_TimeIdlePhysical ignores mouse clicks/mouse moves

GLOBAL CapsLock_VAR_IDLE_1
GLOBAL CapsLock_VAR_IDLE_2
GLOBAL Timer_Delayer
GLOBAL FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING
GLOBAL OLD_CAPS_TOGGLE_STATE
GLOBAL CAPS_DOWN_TRIGGER
GLOBAL SHIFT_CAPS_DOWN_TELL_TRIGGER

GLOBAL OLD_isFullScreen
GLOBAL OSVER_N_VAR

OLD_isFullScreen=0
CapsLock_VAR_IDLE_1=FALSE
CapsLock_VAR_IDLE_2=0
FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
OLD_CAPS_TOGGLE_STATE=-2
CAPS_DOWN_TRIGGER=
SHIFT_CAPS_DOWN_TELL_TRIGGER=FALSE

; -------------------------------------------------------------------
; Set the Time Delay Reminder
; -------------------------------------------------------------------
Timer_Delayer=10000
; -------------------------------------------------------------------


; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6
	

; THE WIDTH IS MEASURE AWAY FROM LEFT OF SCREEN AS SAME X Y
; -------------------------------------------------------------------
; LOWER NUMBER MAKES TIGHTER TO EDGE OF SCREEN
; -------------------------------------------------------------------
width := A_ScreenWidth - 150

; LOWER NUMBER MAKES LOWER TO BOTTOM OF SCREEN
; -------------------------------------------------------------------
height := A_ScreenHeight - 50

; W_P_B = WIDTH_PROGRESS_BAR
; -------------------------------------------------------------------
W_P_B=W69

; F_S_P_B = is the Font Size FS8
; -------------------------------------------------------------------
F_S_P_B=FS8

; H_P_B = HEIGHT_PROGRESS_BAR
; -------------------------------------------------------------------
H_P_B=H20

If OSVER_N_VAR=5
{
	W_P_B=W62
	F_S_P_B=FS6
	H_P_B=H18
	
	; THE WIDTH IS MEASURE AWAY FROM LEFT OF SCREEN AS SAME X Y
	; LOWER NUMBER MAKES TIGHTER TO EDGE OF SCREEN
	width := A_ScreenWidth - 64
	; LOWER NUMBER MAKES LOWER TO BOTTOM OF SCREEN
	height := A_ScreenHeight - 84
}


If OSVER_N_VAR=6
{
	W_P_B=W62
	F_S_P_B=FS6
	H_P_B=H18
	; LOWER NUMBER MAKES TIGHTER TO EDGE OF SCREEN
	width := A_ScreenWidth - 85
	; LOWER NUMBER MAKES LOWER TO BOTTOM OF SCREEN
	height := A_ScreenHeight - 28
}


If OSVER_N_VAR=10
{
	W_P_B=W69
	F_S_P_B=FS8
	H_P_B=H20
	; LOWER NUMBER MAKES TIGHTER TO EDGE OF SCREEN
	width := A_ScreenWidth - 100
	; LOWER NUMBER MAKES LOWER TO BOTTOM OF SCREEN
	height := A_ScreenHeight - (32+4+1)
}
	
	
	
; Gui, Add, Progress, Vertical vMyProgress
; Gui, Show

; OnMessage(0x201, "WM_LBUTTONDOWN")

GOSUB ~*CapsLock

SETTIMER CapsLock_SUB_TIMER,200
SETTIMER CapsLock_SUB_TIMER_1_SECOND,2000
SETTIMER SET_OWN_SCRIPT_LESS_PRIORITY_DEPEND_COMPUTER_NAME, 2000 

RETURN


SET_OWN_SCRIPT_LESS_PRIORITY_DEPEND_COMPUTER_NAME:
	
	; SETTIMER SET_OWN_SCRIPT_LESS_PRIORITY_DEPEND_COMPUTER_NAME, 2000 
	SETTIMER SET_OWN_SCRIPT_LESS_PRIORITY_DEPEND_COMPUTER_NAME, 600000 ; 10 MINUTE
	
	SET_GO_COMPUTERNAME_02=0
	IF (A_ComputerName = "1-ASUS-X5DIJ") 
		SET_GO_COMPUTERNAME_02=1
	IF (A_ComputerName = "2-ASUS-EEE") 
		SET_GO_COMPUTERNAME_02=1
	IF (A_ComputerName = "3-LINDA-PC") 
		SET_GO_COMPUTERNAME_02=1
	
	IF SET_GO_COMPUTERNAME_02=1
		PROCESS, Priority, % DllCall("GetCurrentProcessId"), Low
RETURN




; ----
; Click On Progress Bar? - AutoHotkey Community
; https://www.autohotkey.com/boards/viewtopic.php?t=66561
; --------
; OnMessage() - Syntax & Usage | AutoHotkey
; https://www.autohotkey.com/docs/commands/OnMessage.htm
; --------
; Click On Progress Bar? - AutoHotkey Community
; https://www.autohotkey.com/boards/viewtopic.php?t=66561
; --------
; ProgressBar Examples - AutoHotkey Community
; https://www.autohotkey.com/boards/viewtopic.php?t=4662
; --------
; Clickable ProgressBar, subclassed to ListView/StatusBar - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/70135-clickable-progressbar-subclassed-to-listviewstatusbar/
; ----

; ----
; OnMessage() - Syntax & Usage | AutoHotkey
; https://www.autohotkey.com/docs/commands/OnMessage.htm
; ----



; WM_LBUTTONDOWN(wParam, lParam)
; {
    ; X := lParam & 0xFFFF
    ; Y := lParam >> 16
    ; if A_GuiControl
        ; Control := "`n(in control " . A_GuiControl . ")"
    ; ; ToolTip You left-clicked in Gui window #%A_Gui% at client coordinates %X%x%Y%.%Control%
	; MSGBOX You left-clicked in Gui window #%A_Gui% at client coordinates %X%x%Y%.%Control%
; }

; OnMessage(0x201,"WM_LBUTTONDOWN")
; WM_LBUTTONDOWN() {
; MSGBOX % A_GuiControl
	; If (A_GuiControl = "PrBar1") {
      ; Msgbox % "You've clicked on the progress bar"
   ; }
; }

; -------------------------------------------------------------------
; END OF INIT
; CODE ROUTINES BEGIN
; -------------------------------------------------------------------

CapsLock_SUB_TIMER_1_SECOND:

	isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
	isFullScreen_VAR:=isFullScreen
	if isFullScreen_VAR<1
		isFullScreen_VAR=0

	if isFullScreen_VAR<>%OLD_isFullScreen%
		if isFullScreen_VAR<1
			IF GetKeyState("CapsLock", "T")
			{
				; -----------------------------------------------------------
				; %F_S_P_B% __ is the Font Size FS8
				; %W_P_B% Width of Box 
				; H24 Height
				; -----------------------------------------------------------
				Progress, B1 %W_P_B% %H_P_B% ZH0 %F_S_P_B% WS900 x%width% y%height% CTFF0000, CAPS ON
				Title_Script := " - AutoHotkey v" A_AhkVersion
				WinSet, Top,, % A_ScriptFullPath . Title_Script
				SOUNDBEEP 3500,100
				SOUNDBEEP 2500,100
			}
	OLD_isFullScreen:=isFullScreen_VAR

RETURN


CapsLock_SUB_TIMER:

	if A_PriorKey=CapsLock
		RETURN
	if A_PriorKey=RShift
		RETURN
	if A_PriorKey=LShift
		RETURN

	IF FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=TRUE
	{
		
		IF A_TimeIdlePhysical<%CapsLock_VAR_IDLE_2%
		{
			FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
			IF GetKeyState("CapsLock", "T")
			{
				SOUNDBEEP 3000,200
			}
			ELSE
			{
				SOUNDBEEP 1000,200
			}
		}
	}

	
	IF A_TimeIdlePhysical > %Timer_Delayer%
	{
		IF CapsLock_VAR_IDLE_1=FALSE
		{
		isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
		if isFullScreen
			{
				Progress, off
			}
			ELSE
			{
				Progress, B1 %W_P_B% %H_P_B% ZH0 %F_S_P_B% WS900 x%width% y%height% CTFF0000, CAPS ON
				Title_Script := " - AutoHotkey v" A_AhkVersion
				WinSet, Top,, % A_ScriptFullPath . Title_Script
			}
		}
		CapsLock_VAR_IDLE_1=TRUE
		SHIFT_CAPS_DOWN_TELL_TRIGGER=TRUE
	}
		
	IF CapsLock_VAR_IDLE_2>%A_TimeIdlePhysical%
	{
		IF CapsLock_VAR_IDLE_1=TRUE
		{
			if GetKeyState("CapsLock", "T")
			{
				; -----------------------------------------------------------
				; FS8 __ is the Font Size 
				; %W_P_B% Width of Box and H24 Height
				; -----------------------------------------------------------
				isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
				if isFullScreen
					{
						Progress, off
					}
					ELSE
					{
						Progress, B1 %W_P_B% %H_P_B% ZH0 %F_S_P_B% WS900 x%width% y%height% CTFF0000, CAPS ON
						Title_Script := " - AutoHotkey v" A_AhkVersion
						WinSet, Top,, % A_ScriptFullPath . Title_Script
					}
				SOUNDBEEP 3000,200
			}
			ELSE
			{
				SOUNDBEEP 1000,200
			}
		}
		CapsLock_VAR_IDLE_1=FALSE
	}
		
	CapsLock_VAR_IDLE_2=%A_TimeIdlePhysical%

RETURN
	
Shift::
	if GetKeyState("CapsLock", "T")
	{
		; -----------------------------------------------------------
		; FS8 __ is the Font Size 
		; %W_P_B% Width of Box and H24 Height
		; -----------------------------------------------------------
		isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
		if isFullScreen
		{
			Progress, off
		}
		ELSE
		{
			Progress, B1 %W_P_B% %H_P_B% ZH0 %F_S_P_B% WS900 x%width% y%height% CTFF0000, CAPS ON
			Title_Script := " - AutoHotkey v" A_AhkVersion
			WinSet, Top,, % A_ScriptFullPath . Title_Script
		}
		IF SHIFT_CAPS_DOWN_TELL_TRIGGER=TRUE
		{
			SHIFT_CAPS_DOWN_TELL_TRIGGER=FALSE
			SOUNDBEEP 3000,200
		}
	}
	ELSE
	{
		IF SHIFT_CAPS_DOWN_TELL_TRIGGER=TRUE
		{
			SHIFT_CAPS_DOWN_TELL_TRIGGER=FALSE
		 	SOUNDBEEP 1000,200
		}
	}

RETURN

; -------------------------------------------------------------------
; ENTRY 001 WEB PAGE SOURCE _ CREDIT GIVEN
; -------------------------------------------------------------------


~*CapsLock::
	
	CAPS_TOGGLE_STATE:=GetKeyState("CapsLock", "T")
	
	IF CAPS_TOGGLE_STATE=%OLD_CAPS_TOGGLE_STATE%
	{
		RETURN 
	}
	OLD_CAPS_TOGGLE_STATE:=%CAPS_TOGGLE_STATE%


	if CAPS_TOGGLE_STATE=1
	{
		; -----------------------------------------------------------
		; FS8 __ is the Font Size 
		; %W_P_B% Width of Box and H24 Height
		; -----------------------------------------------------------
		isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
		if isFullScreen
		{
			Progress, off
		}
		ELSE
		{
			Progress, B1 %W_P_B% %H_P_B% ZH0 %F_S_P_B% WS900 x%width% y%height% CTFF0000, CAPS ON
		}
		SOUNDBEEP 3000,200

	}
	else
	{
		SOUNDBEEP 1000,200
		Progress, off
	}
	

	; WHEN CAPS LOCK IS PRESS IT MAKE A SOUND TO INDICATE HIGH OR LOW
	; AND ALSO WHEN BEEN AWAY FROM KEYBOARD LIKE 20 SECOND THAT HAPPEN SOUND ALSO
	; AND WE WANT AFTER CAPS LOCK PRESS AND INDICATE ALSO THE NEXT KEY PRESS AFTER ALSO HAS SOUND INDICATE
	; SET A FLAG TURN SPEED OF TIMER UP A GEAR
	FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=true

	CAPS_DOWN_TRIGGER=TRUE
	SHIFT_CAPS_DOWN_TELL_TRIGGER=TRUE
	
	 
RETURN



; -------------------------------------------------------------------
; ENTRY 002 WEB PAGE SOURCE
; -------------------------------------------------------------------

;=============================================================================================
; Show a ToolTip that shows the current state of the lock keys (e.g. CapsLock) when one is pressed
;=============================================================================================
; ~*NumLock::
; ~*CapsLock::
; ~*ScrollLock::
; Sleep, 10	; drastically improves reliability on slower systems (took a loooong time to figure this out)

; msg := "Caps Lock: " (GetKeyState("CapsLock", "T") ? "ON" : "OFF") "`n"
; msg := msg "Num Lock: " (GetKeyState("NumLock", "T") ? "ON" : "OFF") "`n"
; msg := msg "Scroll Lock: " (GetKeyState("ScrollLock", "T") ? "ON" : "OFF")

; ToolTip, %msg%
; Sleep, 250	; SPECIFY DISPLAY TIME (ms)
; ToolTip		; remove
; return



; ------------------------------------------------------------------
isWindowFullScreen( winTitle ) {
; checks if the specified window is full screen

winID := WinExist( winTitle )
If ( !winID )
	Return false

; ONLY CHECK VALID TITLE WITH TEXT OR ELSE AFRAID CAPTURE DESKTOP 
WinGetTitle, Title_VAR, ahk_id %winID%
If ( !Title_VAR )
	Return false
	
; MSGBOX % Title_VAR
; WINDOWS XP REPORT Program Manager FOR DESKTOP
; ---------------------------------------------
IF Title_VAR=Program Manager	
	Return false
	
WinGet style, Style, ahk_id %WinID%
WinGetPos ,,,winW,winH, ahk_id %WinID%

; 0x800000 is WS_BORDER.
; 0x20000000 is WS_MINIMIZE.
; no border and not minimized

; Return (style & 0x20800000) ? false : true

Return ((style & 0x20800000) 
or winH < A_ScreenHeight 
or winW < A_ScreenWidth) ? false : true

; ----
; Detect Fullscreen application? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/38882-detect-fullscreen-application/
; ----
}





#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk

 
;# ------------------------------------------------------------------
; END BLOCK OF CODE -- EXIT ROUTINE
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
	return (match > 1)
	}
Return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
; -------------------------------------------------------------------

; -------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
; -------------------------------------------------------------------
; EXIT THE APP
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; -------------------------------------------------------------------

; 1 OF 5 EXAMPLE THIS WORKS
IF A_ComputerName=7-ASUS-GL522VW
	SET_EXIT=TRUE
; 2 OF 5 EXAMPLE THIS WORKS
IF (A_ComputerName=7-ASUS-GL522VW)
	SET_EXIT=TRUE
; 3 OF 5 EXAMPLE AND THIS WORKS
IF (A_ComputerName="7-ASUS-GL522VW")
	SET_EXIT=TRUE
; 4 OF 5 EXAMPLE AND THIS NOT WORK
IF A_ComputerName="7-ASUS-GL522VW"
	SET_EXIT=TRUE
; 5 OF 5 EXAMPLE SOMETIMES THIS NOT WORK AND INSTEAD (A_ComputerName="7-ASUS-GL522VW")
IF A_ComputerName=7-ASUS-GL522VW
	SET_EXIT=TRUE

; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; REFERENCE CREDIT
; GET LAST KEY
; -------------------------------------------------------------------
; Getting the last key that was pressed - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/93659-getting-the-last-key-that-was-pressed/
; -------------------------------------------------------------------
; -------------------------------------------------------------------
