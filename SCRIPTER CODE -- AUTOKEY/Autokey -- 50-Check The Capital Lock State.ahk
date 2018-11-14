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
;# __ TO   Wed 31-Oct-2018 01:25:00
;# __ TO   Wed 31-Oct-2018 02:41:00 __ EXTENDED PLAYER 
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
; ---- LOCATION ON-LINE
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
; Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
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

#InstallKeybdHook  ;A_TimeIdlePhysical ignores mouse clicks/mouse moves

GLOBAL CapsLock_VAR_IDLE_1
GLOBAL CapsLock_VAR_IDLE_2
GLOBAL Timer_Delayer
GLOBAL FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING
GLOBAL COUNTER
GLOBAL OLD_CAPS_TOGGLE_STATE
CapsLock_VAR_IDLE_1=FALSE
CapsLock_VAR_IDLE_2=0
FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
COUNTER=0

; -------------------------------------------------------------------
; Set the Time Delay Reminder
; -------------------------------------------------------------------
Timer_Delayer=10000
; -------------------------------------------------------------------

width := A_ScreenWidth - 150
height := A_ScreenHeight - 50

GOSUB ~*CapsLock

SETTIMER CapsLock_SUB_TIMER,50

RETURN

; -------------------------------------------------------------------
; END OF INIT
; CODE ROUTINES BEGIN
; -------------------------------------------------------------------

CapsLock_SUB_TIMER:

	; TOOLTIP % FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING

	tooltip % A_TimeIdlePhysical
	
	; IF A_TimeIdlePhysical>1000
		; TOOLTIP 

	IF A_TimeIdlePhysical=0
		RETURN


	GetKeyState, state, CapsLock
	if state = D 	
	{
		; FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=false
		; CapsLock_VAR_IDLE_2=0
		RETURN
	}
	
	IF FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=TRUE
	{
			
		IF A_TimeIdlePhysical<%CapsLock_VAR_IDLE_2%
		{
			; TOOLTIP "HERE 02"
			FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
			IF GetKeyState("CapsLock", "T")
			{
				SOUNDBEEP 3000,200
		TOOLTIP % COUNTER
		COUNTER+=1
			}
			ELSE
			{
				SOUNDBEEP 1000,200
		TOOLTIP % COUNTER
		COUNTER+=1
			}
		}
	}

	IF A_TimeIdlePhysical > %Timer_Delayer%
	{
		IF CapsLock_VAR_IDLE_1=FALSE
		{
			Progress, B1 W108 H24 ZH0 FS8 WS900 x%width% y%height% CTFF0000, CAPS LOCK ON
		}
		CapsLock_VAR_IDLE_1=TRUE
	}
		
	IF CapsLock_VAR_IDLE_2>%A_TimeIdlePhysical%
	{
		IF CapsLock_VAR_IDLE_1=TRUE
		{
			if GetKeyState("CapsLock", "T")
			{
				; -----------------------------------------------------------
				; FS8 __ is the Font Size 
				; W108 Width of Box and H24 Height
				; -----------------------------------------------------------
				Progress, B1 W108 H24 ZH0 FS8 WS900 x%width% y%height% CTFF0000, CAPS LOCK ON
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
		; W108 Width of Box and H24 Height
		; -----------------------------------------------------------
		Progress, B1 W108 H24 ZH0 FS8 WS900 x%width% y%height% CTFF0000, CAPS LOCK ON
		SOUNDBEEP 3000,50
	}
	ELSE
	{
		; 	SOUNDBEEP 1000,50
	}

RETURN

; -------------------------------------------------------------------
; ENTRY 001 WEB PAGE SOURCE _ CREDIT GIVEN
; -------------------------------------------------------------------


~*CapsLock::
	
	CAPS_TOGGLE_STATE:=GetKeyState("CapsLock", "T")

	GetKeyState, state, CapsLock
	if state = U
		msgbox "22"

	if state = U
		RETURN
	
	IF CAPS_TOGGLE_STATE=%OLD_CAPS_TOGGLE_STATE%
	{
		; FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
		; CapsLock_VAR_IDLE_2=0

		RETURN 
	}
	if CAPS_TOGGLE_STATE=1
	{
		; -----------------------------------------------------------
		; FS8 __ is the Font Size 
		; W108 Width of Box and H24 Height
		; -----------------------------------------------------------
		Progress, B1 W108 H24 ZH0 FS8 WS900 x%width% y%height% CTFF0000, CAPS LOCK ON
		SOUNDBEEP 3000,200
		; TOOLTIP % COUNTER
		; COUNTER+=1

	}
	else
	{
		SOUNDBEEP 1000,200
		Progress, off
		; TOOLTIP % COUNTER
		; COUNTER+=1
	}

	
	OLD_CAPS_TOGGLE_STATE:=CAPS_TOGGLE_STATE

	FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=FALSE
	CapsLock_VAR_IDLE_2=0
	 
RETURN

CapsLock up::
	; WHEN CAPS LOCK IS PRESS IT MAKE A SOUND TO INDICATE HIGH OR LOW
	; AND ALSO WHEN BEEN AWAY FROM KEYBOARD LIKE 20 SECOND THAT HAPPEN SOUND ALSO
	; AND WE WANT AFTER CAPS LOCK PRESS AND INDICATE ALSO THE NEXT KEY PRESS AFTER ALSO HAS SOUND INDICATE
	; SET A FLAG TURN SPEED OF TIMER UP A GEAR
	GetKeyState, state, CapsLock
	if state = D
		RETURN
		
	TOOLTIP "HERE 20"
	FLAG_KEYBOARD_WANT_ANOTHER_SOUNDBEEPER_GOING=TRUE
	CapsLock_VAR_IDLE_2=0
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



 
;# ------------------------------------------------------------------
; USUAL END BLOCK OF CODE TO HELP EXIT ROUTINE
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
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
; -------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

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
; exit the app
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
