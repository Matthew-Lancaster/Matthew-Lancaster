;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 50-Check The Capital Lock State.ahk
;# __ 
;# __ Autokey -- 50-Check The Capital Lock State.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ FROM Tue 30-Oct-2018 22:07:50 __ CODE KNOCK UP TIME
;# __ TO   Wed 31-Oct-2018 00:48:00
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; ---- LOCATION ONLINE
; -------------------------------------------------------------------
; ---- GITHUB.COM
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; This Code Is Duplicated Coder Located at 
; Because I Already Have Enough AutoHotKeys Running at Moment
; Nice to Keep it Separate
; -------------------------------------------------------------------
; Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; DESCRIPTION
;# ------------------------------------------------------------------
; AN INDICATOR FOR CAPS LOCK IS ON
; MY METHOD INTRO LEACHED CREDIT GIVEN SOURCE FROM
; HAS MY EXTRA SKILLFUL THOUGHT PUT IN
; THIS WILL DISPLAY A PROGRESS BAR BOTTOM RIGHT CORNER WITH INFO CAPS 
; LOCK IF IT ON
; IT WILL SOUNDER AT CHANGE OF CAPS LOCK EITHER HI FOR UP AND LOW DOWN
; ALSO A SPECIAL BIT THAT OFTEN UPSET PEOPLE 
; WHEN RETURN TO KEYBOARD AFTER A MINUTE OR 20 SECONDS AWAY AND BEGIN KEY-ER
; IT WILL NOTIFY WITH SOUND BEEPER THAN ALSO HI OR LOW TONE
; I LIKED IT THAT WITH LOGITECH MOUSE 
; BUT THE SOFTWARE WORKED AT FIRST INSTALL AND NEVER AGAIN 
; THE HI LOW SOUND BUT MADE A ANNOYING BEEP INSTEAD
; IT ANNOYING ALSO WITH LOGITECH THE MOUSE DISPLAY OF AN ON SCREEN BOX AT CAPS CHANGE
; CAN;T TURN IT OFF ANYMORE
; AND DOUBLE ANNOYING THE LOGITECH OPTIONS NEW VERSION OF SOFTWARE HAS ONE AS WELL
; MAKING TWO ALMOST AS IN CONFLICT
; WHEN YOU HAVE BOTH TYPE OF MOUSE AND THEIR SOFTWARE RUNNER
; -------------------------------------------------------------------
; Could Be Added to Press Shift on Own to Turn Caps Lock Off
; Test it Out Later
;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; FROM Tue 30-Oct-2018 22:07:50 __ CODE KNOCK UP TIME
; TO   Wed 31-Oct-2018 00:48:00
;# ------------------------------------------------------------------


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

; ----
; Another On-Screen Caps Lock Indicator - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/100990-another-on-screen-caps-lock-indicator/
; ----


GOSUB ~*CapsLock
GLOBAL CapsLock_VAR_IDLE_1
GLOBAL CapsLock_VAR_IDLE_2
CapsLock_VAR_IDLE_1=FALSE
CapsLock_VAR_IDLE_2=0
SETTIMER CapsLock_SUB_TIMER,100

RETURN

CapsLock_SUB_TIMER:

IF (A_TimeIdle > 20000)
	CapsLock_VAR_IDLE_1=TRUE
	
IF CapsLock_VAR_IDLE_2>%A_TimeIdle%
{
	IF CapsLock_VAR_IDLE_1=TRUE
	if GetKeyState("CapsLock", "T")
	{
		SOUNDBEEP 3000,400
	}
	ELSE
	{
		SOUNDBEEP 1000,400
	}
	CapsLock_VAR_IDLE_1=FALSE
}
	
CapsLock_VAR_IDLE_2=%A_TimeIdle%

RETURN
	

; ENTRY 001 WEB PAGE SOURCE

~*CapsLock::
 
	width := A_ScreenWidth - 150
	height := A_ScreenHeight - 50
	 
	Sleep, 100
	if GetKeyState("CapsLock", "T")
	{
		Progress, B1 W108 H24 ZH0 FS8 WS900 x%width% y%height% CTFF0000, CAPS LOCK ON
		SOUNDBEEP 3000,50
	}
	else
	{
		SOUNDBEEP 1000,50
		Progress, off
	}
	 
return


; ENTRY 002 WEB PAGE SOURCE

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