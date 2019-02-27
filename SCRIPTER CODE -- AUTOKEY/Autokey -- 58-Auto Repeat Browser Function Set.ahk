	 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 58-Auto Repeat Browser Function Set.ahk
;# __ 
;# __ Autokey -- 58-Auto Repeat Browser Function Set.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; WORK DO-ER
; -------------------------------------------------------------------
; SPLIT THE CODE AWAY FROM 
; Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
; -------------------------------------------------------------------
; AND MOVED INTO THIS ONE
; Autokey -- 58-Auto Repeat Play Video Facebook.ahk
; -------------------------------------------------------------------
; 01
; AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO:
; 02
; AUTO_RELOAD_RAIN_ALARM:
; 03 
; AUTO_RELOAD_FACEBOOK_QUICK_SUB:
; 04 
; AUTO_RELOAD_FACEBOOK:
; 
; -------------------------------------------------------------------
; FROM __ Tue 26-Feb-2019 15:56:21
; TO   __ Tue 26-Feb-2019 21:14:00 __ FIVE & HALF HOUR _ TIMING'S HARD AND GO FIGURE IT ALL
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 02 03 
; -------------------------------------------------------------------
; WAKE NEXT DAY AND MORE WORK
; FINE TUNER THE TIMING OF FACEBOOK VIDEO REPEATER
; WHILE DO OTHER THING CATCH UP SAME TIME
; GOT JOB DONE
; NOW LOOK TIME _ DO THE FACEBOOK VIDEO ENDER
; BEEN RUN GOOD NOW 
; BUT ONE THING LEARN SOMETIME THE PAGE IN BROWSER OF VIDEO FACEBOOK COULD DO WITH A REFRESH
; NOT SURE HOW IMPLEMENT THE EXTRA YET SO IF SEE VIDEO NOT START WHEN PRESS SPACE-BAR GOES ON
; AND TEST YOURSELF MANUALLY AND NONE RESPONSE THERE REFRESH THE PAGE
; WHY MY LOW END COMPUTER RUN BETTER AT THIS ONE
; HARD TO IMPLEMENT METHOD AS CODE HAS NONE WAY OF TELL A VIDEO PLAYER GO OR NOT
; DON'T WANT REFRESH WHILE PLAY EVER _ OR STEAL FOCUS WHEN OPERATE ANOTHER THING
; -------------------------------------------------------------------
; TO   __  Wed 27-Feb-2019 14:54:59
; -------------------------------------------------------------------
; THAT WAS DO EXTRA MOD FOR FACEBOOK GENERAL RE-LOADER TO SHOW PRESENCE AND NOTIFICATION
; -------------------------------------------------------------------
; TO   __  Wed 27-Feb-2019 18:16:38
; -------------------------------------------------------------------




;# ------------------------------------------------------------------
; Location On-Line
;--------------------------------------------------------------------
; ----
; Autokey -- 58-Auto Repeat Browser Function Set.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2058-Auto%20Repeat%20Browser%20Function%20Set.ahk
; ----
;# ------------------------------------------------------------------

;
; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

IF A_ComputerName="2-ASUS-EEE"
	Exitapp

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path

GLOBAL XPOS
GLOBAL YPOS

GLOBAL O_ID
O_ID=0

GLOBAL OutputVar_4

GLOBAL O_OutputVar_store
O_OutputVar_store=

GLOBAL PART_RENAME_VAR

GLOBAL STATE_XYPOSCOUNTER
STATE_XYPOSCOUNTER=0

GLOBAL OLD_STATE_XYPOSCOUNTER
OLD_STATE_XYPOSCOUNTER=0

GLOBAL OSVER_N_VAR

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

SET_GO=FALSE
IF A_ComputerName=1-ASUS-X5DIJ
	SET_GO=TRUE
IF A_ComputerName=2-ASUS-EEE
	SET_GO=TRUE

; 01 OF 04
IF SET_GO=TRUE 
{
	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 1000
	RETURN
}

AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
OLD_AUTO_RELOAD_FACEBOOK_VAR=0

; 02 OF 04
IF OSVER_N_VAR>5 
{
	SETTIMER AUTO_RELOAD_FACEBOOK,59000
	SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB,1000
}

SET_GO=TRUE
IF A_ComputerName=2-ASUS-EEE
	SET_GO=FALSE
IF A_ComputerName=3-LINDA-PC
	SET_GO=FALSE

; 03 OF 04
IF SET_GO=TRUE 
{
	SETTIMER AUTO_RELOAD_RAIN_ALARM,10000
}
	
AUTO_HITTER_COUNTER_FACEBOOK_COUNTER=0	
AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
; 04 OF 04
IF OSVER_N_VAR>5 
	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO,1000

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.



RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------


AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO:
	; ---------------------------------------------------------------
	; IN ORDER TO USE THIS ROUTINE 
	; FBPURITY MUST BE INSTALLED BECAUSE AT THE END OF A VIDEO IT TRY 
	; TO SHOW NEXT UP VIDEO AND FBP PUTS THAT OFF
	; THAT IS ABOUT ALL
	; ---------------------------------------------------------------
	; OH AND THAT REASON OF FBP IS NOT WORK THE SAME FOR FIREFOX
	; LOADS UP NEXT VIDEO TO PLAY FIREFOX CRAP
	; ---------------------------------------------------------------
	; THERE IS ANOTHER PROBLEM _ TOO MANY CHROME EXTENSION AND AM 
	; TRYING TO FIND ANOTHER ONE WHICH IS STOP PLAY
	; PROCESS OF ELIMINATION TO FIND
	; ---------------------------------------------------------------
	; [ Wednesday 10:42:50 Am_27 February 2019 ]
	; FOUND THE FAULT EXTENSION 
	; YOU NOT HAVE THIS ONE RUNNER _ IT MAKE THE SPACE BAR NOT WORK PROPER TO PLAY VIDEO
	; ----
    ; Social Video Downloader - Chrome Web Store
    ; https://chrome.google.com/webstore/detail/social-video-downloader/kmminjooemmhhbpkbfmjhknffplmjkfi
    ; ----
	; ---------------------------------------------------------------
	
		
	SetTitleMatchMode 2  ; NOT Specify Full path.

	; FORNICATE
	; ----
    ; Deborah Hall
    ; https://www.facebook.com/profile.php?id=100025231092355&eid=ARDGlkY57WIQPW7bfEoyGk0tJwd97KEwCKjLXSlytbeMiWIGJH-oHAMDxaevFqhHrBu5pI1oIxxHqR2h
    ; ----
	; https://www.facebook.com/100025231092355/videos/278770969640604/
	
	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 600000 ; 10 MINUTE
	; SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 300000   ; 5  MINUTE

	; IF !A_ComputerName="8-MSI-GP62M-7RD"
	; 	RETURN

	
	IF A_ComputerName=3-LINDA-PC
		AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
	
	IF A_ComputerName=4-ASUS-GL522VW
		AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
		
	XR_2=0
	IF AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
	{	
		AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=FALSE
		
		XR_3=
		IfWinExist, ahk_class Chrome_WidgetWin_1
			XR_3=Chrome_WidgetWin_1
		IfWinExist, ahk_class MozillaWindowClass
			XR_3=MozillaWindowClass
		
		XR_2=0
		XR_4=
		IfWinExist, Facebook - Google Chrome
		{
			XR_2=1
			XR_4=Facebook - Google Chrome
		}
		IfWinExist, Deborah Hall -
		{
			XR_2=1
			XR_4=Deborah Hall -
		}
		IfWinExist, Facebook - Mozilla Firefox
		{
			XR_2=1
			XR_4=Facebook - Mozilla Firefox
		}
		
		IF XR_2=0
			XR_3=

		IF XR_3
		{	
			WinActivate, ahk_class %XR_3%
			WinWaitActive, ahk_class %XR_3%
			SLEEP 1000
		}
		IF XR_4
		{	
			WinActivate, %XR_4%
			WinWaitActive, %XR_4%
			SLEEP 1000
		}
	}
	
	XR_3=
	IfWinExist, ahk_class Chrome_WidgetWin_1
		XR_3=Chrome_WidgetWin_1
	IfWinExist, ahk_class MozillaWindowClass
		XR_3=MozillaWindowClass

	IF !XR_3
		RETURN
		
	WinGetCLASS, CLASS_VAR, A
	WinGetTITLE, TITLE_VAR, A

	XR_1=0
	IF INSTR(CLASS_VAR,"Chrome_WidgetWin_1")
	{
		XR_1=1
		XR_3=Chrome_WidgetWin_1
	}
	IF INSTR(CLASS_VAR,"MozillaWindowClass")
	{
		XR_1=1
		XR_3=MozillaWindowClass
	}
	
	IF INSTR(TITLE_VAR,"Facebook - Google Chrome")
		XR_2=1
	IF INSTR(TITLE_VAR,"Facebook - Mozilla Firefox")
		XR_2=1
	IF INSTR(TITLE_VAR,"Deborah Hall -")
		XR_2=1

	AUTO_HITTER_COUNTER_FACEBOOK_COUNTER+=1
	LOOP_COUNTER=0

	IF XR_1>0
		IF AUTO_HITTER_COUNTER_FACEBOOK_COUNTER>0
		{
			AUTO_HITTER_COUNTER_FACEBOOK_COUNTER=0
			Loop
			{
				LOOP_COUNTER+=1
				; TAB NEXT
				; --------
				Send, ^{Tab}
				Sleep, 2000
				; WinActivate, ahk_class Chrome_WidgetWin_1
				; IfWinNotExist, ahk_class Chrome_WidgetWin_1
				; 	Return
				
				
				WinGetTitle, CurrentWindowTitle, ahk_class Chrome_WidgetWin_1
				SET_GO=FALSE
				IF INSTR(CurrentWindowTitle,"Facebook - Google Chrome")
					SET_GO=TRUE
				IF INSTR(CurrentWindowTitle,"Deborah Hall -")
					SET_GO=TRUE
				WinGetTITLE, CurrentWindowTitle, A
				IF INSTR(CurrentWindowTitle,"Facebook - Google Chrome")
					SET_GO=TRUE
				IF INSTR(CurrentWindowTitle,"Deborah Hall -")
					SET_GO=TRUE
				
				IF SET_GO=TRUE 
				{
					SLEEP 1000
					BREAK
				}
				IF LOOP_COUNTER>20
				{
					XR_1=0
					BREAK
				}
			}
		}

		
	IF XR_1>0
		IF XR_2>0
		{
			SLEEP 1000
			; CoordMode, Mouse, Client 
			SENDINPUT ^{HOME}
			SLEEP 400
			SENDINPUT {TAB}
			SLEEP 400
			SENDINPUT {SPACE}
			SLEEP 400
			MouseMove, 80, 200
			SLEEP 400

			SOUNDBEEP 1000,50
			SOUNDBEEP 1500,50
			SOUNDBEEP 2000,50

			; MouseClick, LEFT, 80, 200
		}

		
RETURN


AUTO_RELOAD_RAIN_ALARM:

	SET_GO=FALSE
	IF A_ComputerName=2-ASUS-EEE
		SET_GO=TRUE

	IF SET_GO=TRUE 
	{
		SETTIMER AUTO_RELOAD_RAIN_ALARM, OFF
		RETURN
	}

	; ---------------------------------------------------------------
	; RAIN ALARM HAS INTRO A NEW THING LIKE A NAG SCREEN
	; THAT IS LEFT RUNNER A LONG TIME IT ASK YOU TO REFRESH THE SCREEN
	; HA HA
	; [ Wednesday 13:55:40 Pm_16 January 2019 ]
	; ---------------------------------------------------------------

	SETTIMER AUTO_RELOAD_RAIN_ALARM,3600000 ; 1 HOUR MAYBE TOO MUCH
	
	; If (A_TimeIdle < 8000)
	; {
		; RETURN
	; }
	
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A

	XR_1=0
	IF INSTR(CLASS,"Chrome_WidgetWin_1")
		XR_1=1
	IF INSTR(CLASS,"MozillaWindowClass")
		XR_1=1

	XR_2=0
	IF INSTR(TITLE_VAR,"Rain Alarm - Google Chrome")
		XR_2=1
	IF INSTR(TITLE_VAR,"Rain Alarm - Mozilla Firefox")
		XR_2=1

	IF XR_1>0
		IF XR_2>0
		{
			SENDINPUT {F5}
			; SOUNDBEEP 1000,50
		}
		
RETURN

AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY:
	AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
RETURN

AUTO_RELOAD_FACEBOOK_QUICK_SUB:
	; ---------------------------------------------------
	; IF IT WASN'T ON FACEBOOK AND THEN IT IS ON THAT 
	; SITE AND THEN HERE IS THE CODE TO REFRESH
	; QUICKER TALKER
	; IF WASN'T ON FACEBOOK AND THEN IT IS AGAIN
	; ---------------------------------------------------
	; TIMER REQUIRED FOR NOT TOO QUICK
	
	WinGetTITLE, TITLE_VAR, A

	AUTO_RELOAD_FACEBOOK_VAR=0

	IF INSTR(TITLE_VAR,"Your Notifications - Google Chrome")
		AUTO_RELOAD_FACEBOOK_VAR=1
	;Facebook | Error - Google Chrome
	IF INSTR(TITLE_VAR,"Facebook | Error - Google Chrome")
		AUTO_RELOAD_FACEBOOK_VAR=1
	IF INSTR(TITLE_VAR,"Privacy error - Google Chrome")
		AUTO_RELOAD_FACEBOOK_VAR=1
	
	; ---------------------------------------------------------------
	; WHEN LOAD THE COLOUR PALETTE IT LIKE A WINDOW OUTSIDE OF CHROME 
	; AND WE ARE CHOOSE IGNORE THIS ONE
	; MAKE THINK IS ONE OF NORMAL ONE
	; YES WORKER FIRST TIME
	; ---------------------------------------------------------------
	IF INSTR(TITLE_VAR,"Colour")
		AUTO_RELOAD_FACEBOOK_VAR=1

		
		
	IF OLD_AUTO_RELOAD_FACEBOOK_VAR<>%AUTO_RELOAD_FACEBOOK_VAR%
	IF AUTO_RELOAD_FACEBOOK_VAR=0
	{
		AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=TRUE
		SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY, 4000
	}
	
		
	IF OLD_AUTO_RELOAD_FACEBOOK_VAR<>%AUTO_RELOAD_FACEBOOK_VAR%
		IF AUTO_RELOAD_FACEBOOK_VAR=1
			IF AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
			{
				SENDINPUT {F5}
				SLEEP 4000
				;WINWAIT Your Notifications - Google Chrome
				AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=TRUE
				SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY, 10000
			}
		
	OLD_AUTO_RELOAD_FACEBOOK_VAR=%AUTO_RELOAD_FACEBOOK_VAR%
		
RETURN
		

AUTO_RELOAD_FACEBOOK:


	If (A_TimeIdle < 8000)
	{
		RETURN
	}
	
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A

	XR_1=0
	IF INSTR(CLASS,"Chrome_WidgetWin_1")
		XR_1=1

	XR_2=0
	IF INSTR(TITLE_VAR,"Your Notifications - Google Chrome")
		XR_2=1
	IF INSTR(TITLE_VAR,"Facebook | Error - Google Chrome")
		XR_2=1
	IF INSTR(TITLE_VAR,"Privacy error - Google Chrome")
		XR_2=1


	IF XR_1>0
		IF XR_2>0
		{
			SENDINPUT {F5}
			; SOUNDBEEP 1000,50
		}
		
RETURN





;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return

;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
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
;# ------------------------------------------------------------------
; exit the app

