	 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 58-Auto Repeat Browser Function Set.ahk
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
; FOR CODE WRITE NEW IN
; AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO:
;
; FROM __ Tue 26-Feb-2019 15:56:21
; TO   __ Tue 26-Feb-2019 21:14:00 __ FIVE & HALF HOUR _ TIMING'S HARD AND GO FIGURE IT ALL
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 002 AND 003 
; -------------------------------------------------------------------
; WAKE NEXT DAY AND MORE WORK _ QUITE A LOT OF TUNER ALL WEEK
; FINE TUNER THE TIMING OF FACEBOOK VIDEO REPEATER
; WHILE DO OTHER THING CATCH UP SAME TIME
; GOT JOB DONE
; NOW LOOK TIME _ DO THE FACEBOOK VIDEO ENDER
; BEEN RUN GOOD NOW 
; BUT ONE THING LEARN SOMETIME THE PAGE IN BROWSER OF VIDEO FACEBOOK COULD DO WITH A REFRESH
; NOT SURE HOW IMPLEMENT THAT EXTRA YET SO IF SEE VIDEO NOT START WHEN PRESS SPACE-BAR GOES ON
; AND TEST YOURSELF MANUALLY AND NONE RESPONSE THERE REFRESH THE PAGE
; WHY MY LOW END COMPUTER RUN BETTER AT THIS ONE _ SORT OF
; HARD TO IMPLEMENT METHOD AS CODE HAS NONE WAY OF TELL A VIDEO PLAYER GO OR NOT
; DON'T WANT REFRESH WHILE PLAY EVER _ OR STEAL FOCUS WHEN OPERATE ANOTHER THING
; -------------------------------------------------------------------
; TO   __  Wed 27-Feb-2019 14:54:59
; -------------------------------------------------------------------
; THAT WAS DO EXTRA MOD FOR FACEBOOK GENERAL RE-LOADER TO SHOW PRESENCE AND NOTIFICATION
; -------------------------------------------------------------------
; TO   __  Wed 27-Feb-2019 18:16:38
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 004 
; -------------------------------------------------------------------
; ALL AT A WORKER STATE NOW
; ARRAY BEEN ADD AND USE BY FUNCTION CALL FUNCTION TO USE IN WHATEVER SUBROUTINE
; LITTLE BUG HERE AND THERE
; TOOK WHILE SORT FACEBOOK VIDEO TAB ALONG AND PLAY AFTER ARRAY WORKER 
; THINGS WERE NOT PROPER SORTED NOW
;
; -------------------------------------------------------------------
; TO   __ Count = 345 -- Thu 06-Jun-2019 21:27:04
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 004 
; ----------------------------------------------------
; TODAY WORK HERE AND ADD NEW ROUTINE WITH AN ARRAY
; TIMER_SET_ARRAY_BROWSER_TAB_CLOSE:
; THAT HAS TWO FUNCTION
; ONE IT SOLE WORK IT TO DELETE ANY UN-WANTER TAB THAT COME UP
; LIKE THE HOME HUB TELL CONNECTION NOT PROPER
; AND LOSE FOCUS OF ONE THAT WAS THERE

; AND DOUBLE
; WHEN A TAB REMOVE AND IT IS 
; "404 Page Not Found | CPC"
; IT THEN REMOVE THAN A CHECK FOR EVEN MORE
; AND ALSO
; DO A CONTROL SHIFT TAB TO GO BACK ONE
; WELL THAT WAS INTENSION BUT NOT REQUIRE
; AFTER A 404 PAGE REMOVE ALSO TWO OTHER TYPE
; IS THERE
; NEW BROWSER START IN BATCH OF 10
; AND NEW PAGE GET PUT FOR ONE THEM
; AND ALSO SOME OTHER PAGE CPC NOT QUITE 404 BUT STILL DO WITH DELETER
; ----------------------------------------------------
; LOT OF EFFORT PLAY AROUND WITH CODE
; NOT BEEN A WHILE CODER
; LOST IT ON MY EYE TO SPOT ERROR RUSTY ABOUT IT
; ----------------------------------------------------
; FROM   __ Mon 05-Aug-2019 13:25:04
; TO     __ Mon 05-Aug-2019 17:30:00
; ----------------------------------------------------

; -------------------------------------------------------------------
; MONITOR OF VIDEO FACEBOOK REPEAT HITT PUMP HER COUNTER 
; -------------------------------------------------------------------
; Deborah Hall
; -------------------------------------------------------------------
; Wed 27-Feb-2019
; 179 Views _ 12 OCT 2018 
; 198 Views _ 17 OCT 2018
; 154 Views _ 09 NOV 2018
; 128 Views _ 18 NOV 2018
; 171 Views _ 29 NOV 2018
; -------------------------------------------------------------------
; Wed 27-Feb-2019
; 181 Views _ 12 OCT 2018 + 2 TWO
; 200 Views _ 17 OCT 2018 + 2 TWO
; 157 Views _ 09 NOV 2018 + 3 THREE
; 130 Views _ 18 NOV 2018 + 2 TWO
; 174 Views _ 29 NOV 2018 + 3 THREE
; -------------------------------------------------------------------
; Thu 28-Feb-2019 NEAR MIDNIGHT OF THU-FRI _ PLAY HUNDRED OF TODAY
; 181 Views _ 12 OCT 2018 + 0 NAUGHT
; 203 Views _ 17 OCT 2018 + 3 THREE
; 159 Views _ 09 NOV 2018 + 2 TWO
; 132 Views _ 18 NOV 2018 + 2 TWO
; 176 Views _ 29 NOV 2018 + 2 TWO
; -------------------------------------------------------------------
; Sat 02-Mar-2019 MIDDAY _ 2 DAY LATER
; 185 Views _ 12 OCT 2018 + 4 FOUR
; 208 Views _ 17 OCT 2018 + 5 FIVE
; 164 Views _ 09 NOV 2018 + 5 FIVE
; 136 Views _ 18 NOV 2018 + 4 FOUR
; 178 Views _ 29 NOV 2018 + 2 TWO
; -------------------------------------------------------------------
; Wed 06-Mar-2019 MIDNIGHT of TUE WED _ 4 DAY LATER
; 216 Views _ 12 OCT 2018 + 31 UP
; 233 Views _ 17 OCT 2018 + 25
; 169 Views _ 09 NOV 2018 + 5 FIVE
; 162 Views _ 18 NOV 2018 + TWENTY SIX 20+SIX
; 188 Views _ 29 NOV 2018 + 10
; -------------------------------------------------------------------
; Thu 07-Mar-2019 NEAR MIDNIGHT _ 1 DAY LATER
; 225 Views _ 12 OCT 2018 + 9 UP
; 234 Views _ 17 OCT 2018 + 1
; 170 Views _ 09 NOV 2018 + 1
; 165 Views _ 18 NOV 2018 + 3
; 194 Views _ 29 NOV 2018 + SIX 
; -------------------------------------------------------------------
; Sat 09-Mar-2019 __ 8 PM _ 2 DAY
; 225 Views _ 12 OCT 2018 + 0   __ So my very lovely and generous brother bought me a guitar to keep
; 234 Views _ 17 OCT 2018 + 0   __ So...I wrote and composed my first ever song! Its called Penguin Lullaby
; 170 Views _ 09 NOV 2018 + 0   __ Time for a happier song - Riptide
; 165 Views _ 18 NOV 2018 + 0   __ Hurt
; 194 Views _ 29 NOV 2018 + 0   __ New song - Standing in the Rain
;  73 Views _ 28 FEB 2019 ___   __ Blue - In honor of Beth __ About a week ago 
; -------------------------------------------------------------------
; Sat 16-Mar-2019 __ 7 DAY WEEK
; 239 Views _ 12 OCT 2018 + 14  __ So my very lovely and generous brother bought me a guitar to keep
; 254 Views _ 17 OCT 2018 + 20  __ So...I wrote and composed my first ever song! Its called Penguin Lullaby
; 178 Views _ 09 NOV 2018 + 8   __ Time for a happier song - Riptide
; 190 Views _ 18 NOV 2018 + 25  __ Hurt
; 218 Views _ 29 NOV 2018 + 24  __ New song - Standing in the Rain
;  96 Views _ 28 FEB 2019 + 23  __ Blue - In honor of Beth __ About a week ago 
; -------------------------------------------------------------------
; Mon 25-Mar-2019 __ 9 DAY
; 290 Views _ 12 OCT 2018 + ++  __ So my very lovely and generous brother bought me a guitar to keep
; 309 Views _ 17 OCT 2018 + ++  __ So...I wrote and composed my first ever song! Its called Penguin Lullaby
; 104 Views _ 01 NOV 2018 +     __ New song that I filmed myself playing ago. The song is called Hospital Bed.
; 225 Views _ 09 NOV 2018 + ++  __ Time for a happier song - Riptide
; 243 Views _ 18 NOV 2018 + ++  __ Hurt
; 261 Views _ 29 NOV 2018 + ++  __ New song - Standing in the Rain
;  71 Views _ 29 DEC 2018 +     __ Angels
; 147 Views _ 28 FEB 2019 + ++  __ Blue - In honor of Beth __ About a week ago 
; 143 Views _ 11 MAR 2019 +     __ For my Papa Bryan Donald Hall-T he City of Chicago.
;  75 Views _ 13 MAR 2019 +     __ Had a really bad lost my leave. Try myself with music. Here is Moonshadow.
;  73 Views _ 22 MAR 2019 +     __ She _ YouTube
;   3 Views _ 22 MAR 2019 +     __ Hallelujah
;   8 Views _ 23 MAR 2019 +     __ In my pyjamas with bed hair, but a lovely song nonetheless. :) Follow the Sun
; -------------------------------------------------------------------
; Tue 30-Apr-2019 __ Lot of DAY -- Been Play Couple Day Now Set Counter
; 351 Views _ 12 OCT 2018 + ++  __ So my very lovely and generous brother bought me a guitar to keep
; 372 Views _ 17 OCT 2018 + ++  __ So...I wrote and composed my first ever song! Its called Penguin Lullaby
; 166 Views _ 01 NOV 2018 + ++  __ New song Filmed myself play a few days ago. The song is called Hospital Bed.
; 273 Views _ 09 NOV 2018 + ++  __ Time for a happier song - Riptide
; 286 Views _ 18 NOV 2018 + ++  __ Hurt
; 330 Views _ 29 NOV 2018 + ++  __ New song - Standing in the Rain
; 121 Views _ 29 DEC 2018 +     __ Angels
; 234 Views _ 28 FEB 2019 + ++  __ Blue - In honor of Beth __ About a week ago 
; 219 Views _ 11 MAR 2019 +     __ For my Papa Bryan Donald Hall-T he City of Chicago.
; 183 Views _ 13 MAR 2019 +     __ Had a really bad lost my leave. Try myself with music. Here is Moonshadow.
;  41 Views _ 22 MAR 2019 +     __ She _ YouTube
;   7 Views _ 22 MAR 2019 +     __ Hallelujah
;  24 Views _ 23 MAR 2019 +     __ In my pyjamas with bed hair, but a lovely song nonetheless. :) Follow the Sun

; -------------------------------------------------------------------
; Sat 01-Jun-2019 __ Begin A Play Session
; 484 Views _ 12 OCT 2018 + __ So my very lovely and generous brother bought me a guitar to keep
; 443 Views _ 17 OCT 2018 + __ So...I wrote and composed my first ever song! Its called Penguin Lullaby
; 250 Views _ 01 NOV 2018 + __ New song Filmed myself play a few days ago. The song is called Hospital Bed.
; 412 Views _ 09 NOV 2018 + __ Time for a happier song - Riptide
; 319 Views _ 18 NOV 2018 + __ Hurt
; 338 Views _ 29 NOV 2018 + __ New song - Standing in the Rain
; 166 Views _ 29 DEC 2018 + __ Angels
; 295 Views _ 28 FEB 2019 + __ Blue - In honor of Beth __ About a week ago 
; 291 Views _ 11 MAR 2019 + __ For my Papa Bryan Donald Hall-T he City of Chicago.
; 225 Views _ 13 MAR 2019 + __ Had a really bad lost my leave. Try myself with music. Here is Moonshadow.
; 121 Views _ 22 MAR 2019 + __ She _ YouTube
;  87 Views _ 22 MAR 2019 + __ Hallelujah
; 101 Views _ 23 MAR 2019 + __ In my pyjamas with bed hair, but a lovely song nonetheless. :) Follow the Sun

; -------------------------------------------------------------------
; WELL MORE CODE HAS TO GO INNER NOW REQUIRE A REFRESH PAGE AFTER EVERY PLAY
; GET IT GOING RESULT BETTER
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
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


DEBBY_HALL_PAUSE=TRUE
DEBBY_HALL_PAUSE=FALSE

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

; -------------------------------------------------------------------
SendLevel 1
; -------------------------------------------------------------------
; IF ANOTHER CODE AHK APP IS USE _ IF A_PriorKey=F5
; AND THEN SENDINPUT IS LOW-LEVEL KEY HOOKER 
; AND SYSTEM DOESN'T SEE THE KEY ALL THE TIME 
; AND THEN THIS VARIABLE DOES NOT SEE THE KEY _ IF A_PriorKey=F5
; UNLESS SendLevel 1 __ ALL OTHER CODE WILL READ THE _ IF A_PriorKey=F5
; NOT ONLY WITHIN ONE FILE PROJECT
; -------------------------------------------------------------------
; [ Saturday 21:11:00 Pm_08 June 2019 ]
; -------------------------------------------------------------------

SETTIMER TIMER_PREVIOUS_INSTANCE,1

; IF A_ComputerName=2-ASUS-EEE
	; Exitapp

#InstallKeybdHook
;#InstallMouseHook

	
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

; -------------------------------------------------------------------
; NICE ARRAY HERE BUT IF SET GLOBAL HERE NOT WORK
; AND IF SET ARRAY WHILE INIT LIKE DURING DECLARE GOING ON DOESN'T WORK ALSO
; THE ARRAY REQUIRE TWO USE IN BY TWO SUBROUTINE
; SO ANSWER IS SETUP THE ARRAY IN ANOTHER NEW SUBROUTINE OF IT OWN
; AND THAT WAY WORKER
; QUITE GOOD THAT ONE 
; NOW YOU KNOW HOW TO SETUP ARRAY FOR BIGGER USER
; 
; OOPS VAR ArrayCount HAD TO BE SET GLOBAL ON FURTHER DEBUG METHOD
; OOPS STILL PROBLEM ARRAY USE AS FUNCTION
; AND THE STRUCTURE ARRAY LINE CHANGE NOW USE :="" WITH QUOTE 
; BUTTER QUOTE ARE NOT INCLUDE SO NONE STRIP THERE
; -------------------------------------------------------------------
; [ Tuesday 19:20:50 Pm_04 June 2019 ]
; -------------------------------------------------------------------
; GLOBAL FN_Array_1
; GLOBAL ArrayCount

; -------------------------------------------------------------------
; FN_Array_1 := SET_THE_ARRAY_FACEBOOK_VIDEO()
; -------------------------------------------------------------------

FN_Array_1 := SET_ARRAY_1()
FN_ARRAY_FB_F5 := SET_ARRAY_FB_F5()

; FN_ARRAY_AUTO_KEY := SET_ARRAY_AUTO_KEY()
	
FN_ARRAY_RAINER_F5 := SET_ARRAY_RAINER_F5()

	
; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

SET_GO=TRUE
IF A_ComputerName=1-ASUS-X5DIJ
	SET_GO=FALSE
IF A_ComputerName=2-ASUS-EEE
	SET_GO=FALSE

; 01 OF 04
IF SET_GO=TRUE 
{
	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 1000
}

AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
OLD_AUTO_RELOAD_FACEBOOK_VAR=0

SET_GO=TRUE
IF A_ComputerName=1-ASUS-X5DIJ
	SET_GO=FALSE

; 02 OF 04
IF SET_GO=TRUE
IF OSVER_N_VAR>0
{
	SETTIMER AUTO_RELOAD_FACEBOOK,60000
	SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB,1000
}
IF SET_GO=TRUE
IF OSVER_N_VAR>0
IF A_ComputerName=3-LINDA-PC
{
	SETTIMER AUTO_RELOAD_FACEBOOK,% -1 * 1000 * 60 * 10 ; After10Minute
	SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB,1000
}


RAIN_ALARM_DO_ONCE=FALSE

SET_GO=TRUE
IF A_ComputerName=2-ASUS-EEE
	SET_GO=FALSE
IF A_ComputerName=3-LINDA-PC
	SET_GO=FALSE

TIMER_AUTO_RELOAD_RAIN_ALARM_1=0
TIMER_AUTO_RELOAD_RAIN_ALARM_2=0
	
; 03 OF 04
IF SET_GO=TRUE 
{
	SETTIMER AUTO_RELOAD_RAIN_ALARM,1000
}
	
AUTO_HITTER_COUNTER_FACEBOOK_COUNTER=0	
AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
; 04 OF 04
IF OSVER_N_VAR>5 
	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO,1000

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.


FN_SET_ARRAY_BROWSER_TAB_CLOSE := SET_ARRAY_BROWSER_TAB_CLOSE()

SET_GO=FALSE
IF A_ComputerName=1-ASUS-X5DIJ
	SET_GO=TRUE
IF A_ComputerName=2-ASUS-EEE
	SET_GO=TRUE
IF A_ComputerName=3-LINDA-PC
	SET_GO=TRUE
	
SET_GO=TRUE

; 03 OF 04
IF SET_GO=TRUE 
{
	SETTIMER TIMER_SET_ARRAY_BROWSER_TAB_CLOSE,1000
}

RETURN



SET_ARRAY_1() {
;	GLOBAL ArrayCount ; DECLARE GLOBAL WITHIN THE FUNCTION HERE 1ST AND BE USER EVERYWHERE
	FN_Array_1 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Facebook - Mozilla Firefox"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Facebook - Google Chrome"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Deborah Hall -"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Dibs Dabs -"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="She - YouTube - Google Chrome"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Follow the Sun - YouTube - Google Chrome"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Hallelujah - YouTube - Google Chrome"
RETURN FN_Array_1
}

SET_ARRAY_FB_F5() {
	SET_ARRAY_FB_F5 := []
	ArrayCount := 0
	ArrayCount += 1
	SET_ARRAY_FB_F5[ArrayCount]:="Your Notifications - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_FB_F5[ArrayCount]:="Facebook | Error - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_FB_F5[ArrayCount]:="Privacy error - Google Chrome"
RETURN SET_ARRAY_FB_F5
}

SET_ARRAY_RAINER_F5() {
	SET_ARRAY_RAINER_F5 := []
	ArrayCount := 0
	ArrayCount += 1
	SET_ARRAY_RAINER_F5[ArrayCount]:="502 Bad Gateway"
	ArrayCount += 1
	SET_ARRAY_RAINER_F5[ArrayCount]:="Rain Alarm - Mozilla Firefox"
	ArrayCount += 1
	SET_ARRAY_RAINER_F5[ArrayCount]:="Rain Alarm - Google Chrome"
RETURN SET_ARRAY_RAINER_F5
}



; Loop % FN_Array_1.MaxIndex()
; {
	; Element := FN_Array_1[A_Index]
	; ; MSGBOX %Element%
	; IfWinExist, %Element%
		; XR_2=1
		; XR_4=%Element%
; }
; FN_ARRAY_AUTO_KEY := SET_ARRAY_AUTO_KEY()
SET_ARRAY_AUTO_KEY() {
	SET_ARRAY_AUTO_KEY := []
	ArrayCount := 0
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Your Notifications - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook | Error - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Privacy error - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook - Mozilla Firefox"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Deborah Hall -"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Dibs Dabs -"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="She - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Follow the Sun - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Hallelujah - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Rain Alarm - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Rain Alarm - Mozilla Firefox"
RETURN SET_ARRAY_AUTO_KEY
}

SET_ARRAY_BROWSER_TAB_CLOSE() {
	SET_ARRAY_BROWSER_TAB_CLOSE := []
	ArrayCount := 0
	ArrayCount += 1
	SET_ARRAY_BROWSER_TAB_CLOSE[ArrayCount]:="BT Smart Hub Manager"
	ArrayCount += 1
	SET_ARRAY_BROWSER_TAB_CLOSE[ArrayCount]:="404 Page Not Found | CPC"
	
	; ArrayCount += 1
	; SET_ARRAY_BROWSER_TAB_CLOSE[ArrayCount]:="Home"
	
RETURN SET_ARRAY_BROWSER_TAB_CLOSE
}

; BT Smart Hub Manager
; ----
; Home
; http://bthomehub.home/00000110500/gui/#/main/
; ----


 

; -------------------------------------------------------------------
; -------------------------------------------------------------------

AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5:
	; -------------------------------------------------------------------
	; AFTER EVERY SONG WAIT AND PRESS F5 REFRESH 
	; MIGHT IMPROVE RESULT AS GOT NONE HITTER LAST LOT
	; -------------------------------------------------------------------

	SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5,OFF

	IF A_ComputerName<>4-ASUS-GL522VW
		RETURN
	
	; If (A_TimeIdle < 10000)
	; {
		; RETURN
	; }

	XR_3=
	IfWinExist, ahk_class Chrome_WidgetWin_1
		XR_3=Chrome_WidgetWin_1
	IfWinExist, ahk_class MozillaWindowClass
		XR_3=MozillaWindowClass
	

	XR_2=0
	XR_4=


	; Loop % ArrayCount
	Loop % FN_Array_1.MaxIndex()
	{
		Element := FN_Array_1[A_Index]
		IfWinExist, %Element%
		{
			XR_2=1
			XR_4=%Element%
		}
	}
	
	IF XR_2=0
		XR_3=

	IF XR_3
		IF XR_4
		{	
			; SOUNDBEEP 2000,100
			WinActivate, %XR_4%
			WinWaitActive
			SLEEP 100
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


	; Loop % ArrayCount
	Loop % FN_Array_1.MaxIndex()
	{
		Element := FN_Array_1[A_Index]
		IF INSTR(TITLE_VAR,Element)
			XR_2=1
	}

	; IF INSTR(TITLE_VAR,"Facebook - Google Chrome")
	; XR_2=1
	; IF INSTR(TITLE_VAR,"Facebook - Mozilla Firefox")
	; XR_2=1
	; IF INSTR(TITLE_VAR,"Deborah Hall -")
	; XR_2=1
	; IF INSTR(TITLE_VAR,"She - YouTube - Google Chrome")
	; XR_2=1
	; IF INSTR(TITLE_VAR,"Follow the Sun - YouTube - Google Chrome")
	; XR_2=1
	; IF INSTR(TITLE_VAR,"Hallelujah - YouTube - Google Chrome")
	; XR_2=1

	AUTO_HITTER_COUNTER_FACEBOOK_COUNTER+=1
	LOOP_COUNTER=0

	IF XR_1>0
		IF XR_2>0
		{
			SLEEP 1000
			SENDINPUT {F5}

			SOUNDBEEP 1500,50
			SOUNDBEEP 2000,50
		}

	; # Win (Windows logo key) 
	; ! Alt 
	; ^ Control 
	; + Shift 
	; & An ampersand may be used between any two keys or mouse buttons to combine them into a custom hotkey. 
		
RETURN

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
	; ALSO WANT THIS ONE IT STOP VIDEO AUTO PLAY WHEN PAGE JUST LOADER
	; DOUBLE CHECK AND NOT ALWAYS REQUIRE THIS AS NOT REQUIRE NOW -- 09 MAR 2019
	; ----
	; Disable HTML5 Autoplay - Chrome Web Store
	; https://chrome.google.com/webstore/detail/disable-html5-autoplay/efdhoaajjjgckpbkoglidkeendpkolai
	; -----
	; [ Saturday 12:54:00 Pm_02 March 2019 ]
	; SEEM FUNNY DIDN'T HITTER ME AT FIRST
	; WHEN TALKER MY FACEBOOK PAGE
	; ABOUT HOW FACEBOOK DESTROY THE REAL HITT COUNTER ONLY ALLOW IT UNIQUE HITTER
	; THE A BIT ANTI-GET YOUR HITTER 
	; AS EXAMPLE YOUTUBE DON'T WOK THAT WAY
	; I WAS TALKER IT PRETTY AMAZING IF GOT MILLIONS HITTER WHEN ALL UNIQUE-ER
	; GOES TO SHOW PEOPLE WATCH ANYTHING WHEN FIRST COME ON SCREEN
	; AND FACEBOOK ALWAYS HAD A HITT POLICY OF THEIR UP FOR THING ZACK AND NO ANYBODY ELSE
	; ---------------------------------------------------------------

	IF A_ComputerName<>4-ASUS-GL522VW
		RETURN
	
	IF DEBBY_HALL_PAUSE=TRUE
		RETURN
	
	SetTitleMatchMode 2  ; NOT Specify Full path.
	
	; FORNICATE PLEASURE
	; ----
    ; Deborah Hall
    ; https://www.facebook.com/profile.php?id=100025231092355&eid=ARDGlkY57WIQPW7bfEoyGk0tJwd97KEwCKjLXSlytbeMiWIGJH-oHAMDxaevFqhHrBu5pI1oIxxHqR2h
    ; ----
	; https://www.facebook.com/100025231092355/videos/278770969640604/
	
	; ----
	; Hi Room I Got 11 Code Project Been Updated This... - Matthew Lancaster
	; https://www.facebook.com/matthew.lancaster.4/posts/10211817982758757
	; ----
	 
	If (A_TimeIdle < 10000)
	{
		RETURN
	}
	 
	IF A_ComputerName=3-LINDA-PC
	{
		; SETTER SO PLAY AGAIN BY ACTIVATE SCREEN TO FORCE IN FOCUS AND PLAY
		; IS OPTION THIS ONLY HAPPEN ON FIRST RUN
		AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
		SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 2400000 ; 40 MINUTE
		; SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 6000000 ; 1 HOUR  
		; SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 1200000 ; 20 MINUTE
	}
	ELSE
	{
		; SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 1200000 ; 20 MINUTE
		SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 600000 ; 10 MINUTE
		; SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, 300000   ; 5  MINUTE
	}

	SET_GO=TRUE
	IF A_ComputerName=5-ASUS-P2520LA
	{
		SETTIMER AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO, OFF
		RETURN
	}

	SET_GO=TRUE
	IF A_ComputerName=7-ASUS-GL522VW
		SET_GO=FALSE
	IF A_ComputerName=8-MSI-GP62M-7RD
		SET_GO=FALSE
		
	IF SET_GO=FALSE
		RETURN

	; IF SET_GO=FALSE
	; {
		; WinGetCLASS, CLASS_VAR, A
		; WinGetTITLE, TITLE_VAR, A

		; XR_1=
		; IF INSTR(CLASS_VAR,"Chrome_WidgetWin_1")
		; {
			; XR_1=1
			; XR_3=Chrome_WidgetWin_1
		; }
		; IF INSTR(CLASS_VAR,"MozillaWindowClass")
		; {
			; XR_1=1
			; XR_3=MozillaWindowClass
		; }
		
		; XR_2=
		; Loop % ArrayCount
		; {
			; Element := FN_Array_1[A_Index]
			; IF INSTR(TITLE_VAR,Element)
				; XR_2=1
		; }
		
		; ; MSGBOX % XR_2

		; ; XR_2=
		; ; IF INSTR(TITLE_VAR,"Facebook - Google Chrome")
			; ; XR_2=1
		; ; IF INSTR(TITLE_VAR,"Facebook - Mozilla Firefox")
			; ; XR_2=1
		; ; IF INSTR(TITLE_VAR,"Deborah Hall -")
			; ; XR_2=1
		; ; IF INSTR(TITLE_VAR,"She - YouTube - Google Chrome")
			; ; XR_2=1
		; ; IF INSTR(TITLE_VAR,"Follow the Sun - YouTube - Google Chrome")
			; ; XR_2=1
		; ; IF INSTR(TITLE_VAR,"Hallelujah - YouTube - Google Chrome")
			; ; XR_2=1
			
		; IF (!XR_1 or !XR_2)
		; {
			; RETURN
		; }
	; }
		
	
	IF A_ComputerName=3-LINDA-PC
	{
		; SETTER SO PLAY AGAIN BY ACTIVATE SCREEN TO FORCE IN FOCUS AND PLAY
		; IS OPTION THIS ONLY HAPPEN ON FIRST RUN
		AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_RUN_ONCE=TRUE
	}
	
	; ---------------------------------------------------------------
	; DODGY LINE LEFT IN 
	; MAKE WORK ALL THE TIME -- WHEN WINDOW NOT IN FOCUS
	; HAVE A TRY
	; ---------------------------------------------------------------
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
		;Loop % ArrayCount
		Loop % FN_Array_1.MaxIndex()
		{
			Element := FN_Array_1[A_Index]
			; MSGBOX %Element%
			IfWinExist, %Element%
				XR_2=1
				XR_4=%Element%
		}
		
		IF XR_2=0
			XR_3=

		IF XR_3
			IF XR_4
			{	
				WinActivate, %XR_4%
				WinWaitActive
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
	
	XR_2=
	; Loop % ArrayCount
	Loop % FN_Array_1.MaxIndex()
	{
		Element := FN_Array_1[A_Index]
		IF INSTR(TITLE_VAR,Element)
			XR_2=1
	}

	AUTO_HITTER_COUNTER_FACEBOOK_COUNTER+=1
	LOOP_COUNTER=0

	IF XR_1>0
	{
		IF XR_2>0
		{
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
					
					WinGetTitle, CurrentWindowTitle, ahk_class Chrome_WidgetWin_1
					SET_GO=FALSE
					
					; Loop % ArrayCount
					Loop % FN_Array_1.MaxIndex()
					{
						Element := FN_Array_1[A_Index]
						IF INSTR(CurrentWindowTitle,Element)
							{
								SET_GO=TRUE
							}
					}
						
					WinGetTITLE, CurrentWindowTitle, A
					; Loop % ArrayCount
					Loop % FN_Array_1.MaxIndex()
					{
						Element := FN_Array_1[A_Index]
						IF INSTR(CurrentWindowTitle,Element)
							{
								SET_GO=TRUE
							}
					}
					
					SET_GO_YOU=FALSE
					IF INSTR(CurrentWindowTitle,"YouTube - Google Chrome")
						SET_GO_YOU=TRUE
					
					
					IF SET_GO=TRUE 
					{
						XR_2=1
						SLEEP 1000
						BREAK
					}
					IF LOOP_COUNTER>10
					{
						XR_1=0
						BREAK
					}
				}
			}
		}
	}
		
	IF XR_1>0
		IF XR_2>0
			IF SET_GO_YOU=FALSE 
			{
				; ---------------------------------------------------
				; REQUIRE EXTRA F5 REFRESH PAGE AS WHEN LEFT FOR A WHILE
				; AND COME BACK ASK TO PLAYER ONLY ABOUT 1- SECONDS 
				; BUT NOT OF REFRESH PAGE BEFORE
				; ---------------------------------------------------
				SLEEP 1000
				SENDINPUT {F5}
				SLEEP 8000
				IF A_ComputerName=3-LINDA-PC
					SLEEP 10000

				SLEEP 1500
				; CoordMode, Mouse, Client 
				SENDINPUT ^{HOME}
				SLEEP 500
				SENDINPUT {TAB}
				SLEEP 500
				SENDINPUT {SPACE}
				SLEEP 500
				MouseMove, 80, 200
				SLEEP 400

				SOUNDBEEP 1000,50
				SOUNDBEEP 1500,50
				SOUNDBEEP 2000,50
				
				SETTIMER,  AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5, OFF
				SETTIMER,  AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5, 420000 ; 7 MINUTE

				; MouseClick, LEFT, 80, 200
			}

	IF XR_1>0
		IF XR_2>0
			IF SET_GO_YOU=TRUE
			{
				SLEEP 1000
				SENDINPUT {F5}
				SLEEP 8000

				SLEEP 1500
				; CoordMode, Mouse, Client 
				SENDINPUT ^{HOME}
				SLEEP 500
				SENDINPUT {SPACE}
				SLEEP 500
				MouseMove, 200, 200
				SLEEP 400

				SOUNDBEEP 1000,50
				SOUNDBEEP 1500,50
				SOUNDBEEP 2000,50
				
				SETTIMER,  AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5, OFF
				SETTIMER,  AUTO_HITTER_COUNTER_FOR_FACEBOOK_VIDEO_PRESS_F5, 420000 ; 7 MINUTE

				; MouseClick, LEFT, 80, 200
			}
			
		
RETURN




TIMER_SET_ARRAY_BROWSER_TAB_CLOSE:

	WinGetCLASS, CLASS_FOCUS, A
	WinGetTITLE, TITLE_VAR_FOCUS, A
	WinGetTITLE, TITLE_CHROME, ahk_class Chrome_WidgetWin_1
	WinGetTITLE, TITLE_MOZILLA, ahk_class MozillaWindowClass

	ARRAY_BROWSER_TAB_CLOSE_SET_GO=
	Loop % FN_SET_ARRAY_BROWSER_TAB_CLOSE.MaxIndex()
	{
		Element := FN_SET_ARRAY_BROWSER_TAB_CLOSE[A_Index]
		IF INSTR(TITLE_VAR,Element)
			ARRAY_BROWSER_TAB_CLOSE_SET_GO=%Element%
		IF INSTR(TITLE_CHROME,Element)
		{
			BROWSER_APP=1
			ARRAY_BROWSER_TAB_CLOSE_SET_GO=%Element%
		}
		IF INSTR(TITLE_MOZILLA,Element)
		{
			BROWSER_APP=2
			ARRAY_BROWSER_TAB_CLOSE_SET_GO=%Element%
		}
	}

	If ARRAY_BROWSER_TAB_CLOSE_SET_GO
	{
		IF BROWSER_APP=1
		{
			WinActivate, ahk_class Chrome_WidgetWin_1
			WinWaitActive
		}
		IF BROWSER_APP=2
		{
			WinActivate, ahk_class MozillaWindowClass
			WinWaitActive
		}
		
		WinGetTITLE, TITLE_VAR_2, A
		IF INSTR(TITLE_VAR_2,ARRAY_BROWSER_TAB_CLOSE_SET_GO)
		{
			WinGet, HWND_2, ID, %ARRAY_BROWSER_TAB_CLOSE_SET_GO%
			IF HWND_2
			{
				Send,,^{w}
				SOUNDBEEP 1000,50
				SLEEP 100
			}
			
			IF INSTR(TITLE_VAR_2,"404 Page Not Found | CPC")
			DONE_ONE_COUNTER=10
			LOOP, 100
			{
				WinGetCLASS, CLASS_FOCUS, A
				BROWSER_APP=
				IF INSTR(CLASS_FOCUS,"Chrome_WidgetWin_1")
					BROWSER_APP=1
				IF INSTR(CLASS_FOCUS,"MozillaWindowClass")
					BROWSER_APP=2

				HWND_2=
				TITLE_VAR=
				IF BROWSER_APP=1
				{
					WinGetTITLE, TITLE_VAR, ahk_class Chrome_WidgetWin_1
					WinGet, HWND_2, ID, ahk_class Chrome_WidgetWin_1
				}
				IF BROWSER_APP=2
				{
					WinGetTITLE, TITLE_VAR, ahk_class MozillaWindowClass
					WinGet, HWND_2, ID, ahk_class MozillaWindowClass
				}

				DONE_ONE=0
				IF HWND_2
				IF INSTR(TITLE_VAR,"CPC - Google Chrome")
				{
					Send,,^{w}
					SLEEP 100
					DONE_ONE=1
					DONE_ONE_COUNTER+=1
				}
				
				IF DONE_ONE=0
				IF INSTR(TITLE_VAR,"New Tab - Google Chrome")
				{
					Send,,^{w}
					SLEEP 100
					DONE_ONE=1
					DONE_ONE_COUNTER+=1
				}
				IF DONE_ONE=0
				IF INSTR(TITLE_VAR,"404 Page Not Found | CPC")
				{
					Send,,^{w}
					SLEEP 100
					DONE_ONE=1
					DONE_ONE_COUNTER+=1
				}
				
				DONE_ONE_COUNTER-=1
				IF DONE_ONE_COUNTER<0
					BREAK
				
				; SLEEP 100

			}
		}
	}

RETURN

SET_RAIN_ALARM_WINDOW_DIMENSION:
		
		; -----------------------------------------------------------
		; BOTH SAME
		; WINDOWS VB6 REPORT SPY AS 1ST ONE
		; AHK SPY REPORT AS 2ND
		; -----------------------------------------------------------
		; TODAY DONE THE COORDINATE FOR HERE
		; WinMove
		; THAT TAKE INTO ACCOUNT THE WIN XP MY COMPUTER DOCKED TO LEFT SCREEN
		; AND TASK BAR TRAY BAR AT BOTTOM OF SCREEN
		; CENTER IT IN THERE
		; [ Saturday 21:14:00 Pm_02 March 2019 ]
		; SPEND 
		; FROM ____ Sat 02-Mar-2019 19:42:43
		; TO   ____ Sat 02-Mar-2019 21:14:02 _ 1 HALF HOUR
		; WAS SOME OTHER MINOR BUG FIND NOT LOOKED OVER
		; AND PLAY AROUND HERE WHILE CODE OPEN
		; -----------------------------------------------------------
		
		WinGet, active_id, ID, ahk_class ReBarWindow321
		if !active_id
			WinGet, active_id, ID, ahk_class BaseBar
		
		if active_id
		{
			WinGetPos, X, Y, Width, Height, ahk_id %active_id%
		}
		
		if Width>0 
		{
			; msgbox % x "x " y "y " width_2 "w " height
			x=%Width%
			; x+=2
			;x+=1
		}

		if !Width
		{
			;x=1
		}
		Width_2=%A_ScreenWidth%
		Width_2-=%X%
		; Width_2-=2

		WinGet, active_id, ID, ahk_class Shell_TrayWnd
		if active_id
		{
			WinGetPos, , Y_4,, Height_4, ahk_id %active_id%
		}

		if Y_4>0 
		{
			;Y_4-=4
			;Y_4-=1
		}

		if !Y_4
		{
			Y_4=%A_ScreenHeight%
		}
		Height_4=%Y_4%
		; Height_4-=12
		; Height_4-=2
		
		; (A_ScreenWidth/2)-(Width/2), (A_ScreenHeight/2)-(Height/2)
		
		WinGetCLASS, CLASS, A
		WinGetTITLE, TITLE_VAR, A

		SET_GO=FALSE
		IF A_ComputerName=1-ASUS-X5DIJ
			SET_GO=TRUE
		IF A_ComputerName=5-ASUS-P2520LA
			SET_GO=TRUE
		IF A_ComputerName=4-ASUS-GL522VW
			SET_GO=TRUE

		IF SET_GO=TRUE
		{
			RAINER_F5_SET_GO=
			Loop % FN_ARRAY_RAINER_F5.MaxIndex()
			{
				Element := FN_ARRAY_RAINER_F5[A_Index]
				IF INSTR(TITLE_VAR,Element)
					RAINER_F5_SET_GO=%Element%
			}

			If RAINER_F5_SET_GO
			{
				WinActivate, %RAINER_F5_SET_GO%
				WinWaitActive
				SLEEP 400
			}
			
			WinMove, %RAINER_F5_SET_GO3%, ,x, -2, Width_2, Height_4
		}

RETURN


AUTO_RELOAD_RAIN_ALARM:

	SET_GO_8=FALSE
	IF A_ComputerName=1-ASUS-X5DIJ
		SET_GO_8=TRUE
	IF A_ComputerName=5-ASUS-P2520LA
		SET_GO_8=TRUE
	IF A_ComputerName=4-ASUS-GL522VW
		SET_GO_8=TRUE


	IF A_ComputerName=2-ASUS-EEE
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

	
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A

	; QUICK ONE IF PROBLEM 
	; ONLY QUICK ON IF ACTIVE WINDOW 
	; NEED CORRECTION
	; --------------------
	Element := "502 Bad Gateway"
	IF INSTR(TITLE_VAR,Element)
	{
		IF TIMER_AUTO_RELOAD_RAIN_ALARM_1=0
		{
			TIMER_AUTO_RELOAD_RAIN_ALARM_1=%A_NOW%
			TIMER_AUTO_RELOAD_RAIN_ALARM_1+=10,seconds
		}
	}
	
	; USUAL ONE
	; ---------
	IF TIMER_AUTO_RELOAD_RAIN_ALARM_2=0
	{
		TIMER_AUTO_RELOAD_RAIN_ALARM_2=%A_NOW%
		TIMER_AUTO_RELOAD_RAIN_ALARM_2+=30,Minutes
		; TIMER_AUTO_RELOAD_RAIN_ALARM_2+=1,Hours
	}	
	
	SET_GO_2=0
	IF TIMER_AUTO_RELOAD_RAIN_ALARM_1>0
		IF TIMER_AUTO_RELOAD_RAIN_ALARM_1<%A_NOW%
		{
			TIMER_AUTO_RELOAD_RAIN_ALARM_1=0
			SET_GO_2=2
		}
	SET_GO_3=0
	IF TIMER_AUTO_RELOAD_RAIN_ALARM_2>0
		IF TIMER_AUTO_RELOAD_RAIN_ALARM_2<%A_NOW%
		{
			TIMER_AUTO_RELOAD_RAIN_ALARM_2=0
			SET_GO_3=3
		}
	
	IF RAIN_ALARM_DO_ONCE=FALSE
	{
		TIMER_AUTO_RELOAD_RAIN_ALARM_2=%A_NOW%
		TIMER_AUTO_RELOAD_RAIN_ALARM_2+=10,Seconds

		RAIN_ALARM_DO_ONCE=TRUE
		GOSUB SET_RAIN_ALARM_WINDOW_DIMENSION
	}


	; IF ALL TIMER TALK DON'T GO 
	; WANT IT BETTER
	; ONE EXTRA TIMER QUICK WHEN SCRIPT FIRST BEGIN
	; USE 2ND TIMER USUAL FOR THAT
	; --------------------------
	IF SET_GO_2=0
		IF SET_GO_3=0
			RETURN
	
	; SETTIMER AUTO_RELOAD_RAIN_ALARM,3600000 ; 1 HOUR MAYBE TOO MUCH
	
	; If (A_TimeIdle < 8000)
		; RETURN
	
		
	; CHECK FOR WIN EXIST AND MAKE ACTIVE
	; -----------------------------------
	; SET_GO_8 IS COMPUTER NAME
	IF SET_GO_8=TRUE
	{
		Loop % FN_ARRAY_RAINER_F5.MaxIndex()
		{
			Element := FN_ARRAY_RAINER_F5[A_Index]
			IfWinExist, %Element%
			{
				WinActivate, %Element%
				
				; WinWaitActive ---- NOT WORK ON WINDOWS 07 IN THE STYLE SHOW NEXT
				; WinWaitActive, %Element%
				
				WinWaitActive
				SLEEP 200
				BREAK
			}

		}
	}

	; SET_GO_8 IS COMPUTER NAME
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A
	IF SET_GO_8=TRUE
	{
		RAINER_F5_SET_GO=
		Loop % FN_ARRAY_RAINER_F5.MaxIndex()
		{
			Element := FN_ARRAY_RAINER_F5[A_Index]
			IF INSTR(TITLE_VAR,Element)
				RAINER_F5_SET_GO=%Element%
		}

		If RAINER_F5_SET_GO
		{
			WinActivate, %RAINER_F5_SET_GO%
			WinWaitActive
			SLEEP 1000
		}
	}
				
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A

	XR_1=0
	IF INSTR(CLASS,"Chrome_WidgetWin_1")
		XR_1=1
	IF INSTR(CLASS,"MozillaWindowClass")
		XR_1=1

	RAINER_F5_SET_GO=
	Loop % FN_ARRAY_RAINER_F5.MaxIndex()
	{
		Element := FN_ARRAY_RAINER_F5[A_Index]
		IF INSTR(TITLE_VAR,Element)
			RAINER_F5_SET_GO=%Element%
	}

	IF XR_1>0
		If RAINER_F5_SET_GO
		{
			SENDINPUT {F5}
			SOUNDBEEP 1000,50
		}
		
RETURN

AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY:
	AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
	SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY,OFF
RETURN

AUTO_RELOAD_FACEBOOK_QUICK_SUB:
	; ---------------------------------------------------
	; WHEN FACEBOOK PAGE GONE WRONG BECAUSE SOME RELOADER 
	; ERROR
	; MADE TITLE NOT AS EXPECTER
	; ---------------------------------------------------
	; IF IT WASN'T ON FACEBOOK AND THEN IT IS ON THAT 
	; SITE AND THEN HERE IS THE CODE TO REFRESH
	; QUICKER TALKER
	; IF WASN'T ON FACEBOOK AND THEN IT IS AGAIN
	; ---------------------------------------------------
	; TIMER REQUIRED FOR NOT TOO QUICK

	SET_GO=FALSE
	IF A_ComputerName=1-ASUS-X5DIJ
		SET_GO=TRUE
	IF A_ComputerName=5-ASUS-P2520LA
		SET_GO=TRUE
	IF SET_GO
		SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB,OFF
	
		
	
	SET_GO_2=FALSE
	IF A_ComputerName=2-ASUS-EEE		
		SET_GO_2=TRUE
	IF A_ComputerName=3-LINDA-PC
		SET_GO_2=TRUE
	IF A_ComputerName=8-MSI-GP62M-7RD
		SET_GO_2=TRUE
		
	IF SET_GO_2=TRUE
	{
		Loop % FN_ARRAY_FB_F5.MaxIndex()
		{
			Element := FN_ARRAY_FB_F5[A_Index]
			IfWinExist, %Element%
			{
				WinActivate, %Element%
				WinWaitActive
				BREAK
			}
		}
	}
	
	WinGetTITLE, TITLE_VAR, A
	; "Your Notifications - Google Chrome"
	AUTO_RELOAD_FACEBOOK_VAR=0
	Loop % FN_ARRAY_FB_F5.MaxIndex()
	{
		Element := FN_ARRAY_FB_F5[A_Index]
		IF INSTR(TITLE_VAR,Element)
		{
			AUTO_RELOAD_FACEBOOK_VAR=1
			BREAK
		}
	}
	
	; ---------------------------------------------------------------
	; WHEN LOAD THE COLOUR PALETTE IT LIKE A WINDOW OUTSIDE OF CHROME 
	; AND WE ARE CHOOSE IGNORE THIS ONE
	; MAKE THINK IS ONE OF NORMAL ONE
	; YES WORKER FIRST TIME
	; ---------------------------------------------------------------
	IF INSTR(TITLE_VAR,"Colour")
	{
		AUTO_RELOAD_FACEBOOK_VAR=0
		RETURN
	}	
		
	; IF OLD_AUTO_RELOAD_FACEBOOK_VAR<>%AUTO_RELOAD_FACEBOOK_VAR%
		; IF AUTO_RELOAD_FACEBOOK_VAR=1
		; {
			; AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=TRUE
			; SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY, 4000
			; ; WHEN PAGE SWAP DELAY A LITTLE BEFORE ACTION F5
		; }
		
	IF OLD_AUTO_RELOAD_FACEBOOK_VAR<>%AUTO_RELOAD_FACEBOOK_VAR%
		IF AUTO_RELOAD_FACEBOOK_VAR=1
			IF AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=FALSE
			{
				; SLEEP 2000
				SENDINPUT {F5}
				; SLEEP 2000
				AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY_VAR=TRUE
				SETTIMER AUTO_RELOAD_FACEBOOK_QUICK_SUB_DELAY, 10000
				; AFTER ISSUE F5 DELAY IS LONGER
			}
		
	OLD_AUTO_RELOAD_FACEBOOK_VAR=%AUTO_RELOAD_FACEBOOK_VAR%
		
RETURN
		

AUTO_RELOAD_FACEBOOK:

	If (A_TimeIdle < 8000)
		RETURN
	
	SET_GO_2=FALSE
	IF A_ComputerName=2-ASUS-EEE		
		SET_GO_2=TRUE
	IF A_ComputerName=3-LINDA-PC
		SET_GO_2=TRUE
		
	IF SET_GO_2=TRUE
	{
		Loop % FN_ARRAY_FB_F5.MaxIndex()
		{
			Element := FN_ARRAY_FB_F5[A_Index]
			IfWinExist, %Element%
			{
				WinActivate, %Element%
				WinWaitActive
				BREAK
			}
		}
	}

	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_VAR, A

	XR_1=0
	IF INSTR(CLASS,"Chrome_WidgetWin_1")
		XR_1=1

	XR_2=0
	Loop % FN_ARRAY_FB_F5.MaxIndex()
	{
		Element := FN_ARRAY_FB_F5[A_Index]
		IF INSTR(TITLE_VAR,Element)
			XR_2=1
	}

	IF XR_1>0
		IF XR_2>0
		{
			; SendLevel 1 ; ---- USED DURING DECLARE STAGE CODE INITIALIZE ABOVE
			SENDINPUT {F5}
			; SOUNDBEEP 1000,50
		}
		
RETURN


MenuHandler:
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02-INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03-INCLUDE MENU 03 of 03.ahk


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
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
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

