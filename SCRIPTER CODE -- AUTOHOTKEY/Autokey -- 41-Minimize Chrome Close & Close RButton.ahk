;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk
;# __ 
;# __ Autokey -- 41-Minimize Chrome Close & Close RButton.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Tue 16-Oct-2018 03:00:00
;# __ LAST EDITOR
;# __ Sat 20-Oct-2018 06:24:00
;# __ Tue 23-Oct-2018 00:08:00 - END TIME 4 HOURS
;# __ 
;  =============================================================

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


;# ------------------------------------------------------------------
; DESCRIPTION
;# ------------------------------------------------------------------
; QUITE SIMPLY FOR CERTAIN PROGRAMS HARDCODE-ED IN SUBROUTINE HERE
; WILL STOP THE CLOSE BUTTON AND INSTEAD MINIMIZE
; AND ALSO CLOSE WHEN RIGHT CLICKER
; ESPECIALLY FOR CHROME NOT LOSE ALL YOUR TAB TO POSSIBLE FAILURE
; THAT SEEM LIKE A RISK SOMETIMES
; -------------------------------------------------------------------
; BASED ON THE WM_NCHITTEST := 0x0084 DETECT CLOSE WINDOW RED CROSS TICKER
;# ------------------------------------------------------------------
; SOURCE
; ----
; Disable the minimize/maximize/close buttons - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/83455-disable-the-minimizemaximizeclose-buttons/
; ----
; NICE IDEA
; AFTER A RECENT FRIGHT CLOSE OF CHROME UNINTENTIONAL BUT OKAY NONE LOSS OF WORKER
; A LITTLE HELP ON SUBJECT
; MOST WINDOW IS ABLE TO SET CLOSE BUTTON DISABLED BUT NOT CHROME
; MAYBE CODED TO SWITCH BACK ON _ INTENTION WAS THERE
; ----
; Close window when mouse hover red X square box? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/81233-close-window-when-mouse-hover-red-x-square-box/
; ----
; ----
; Close open windows when mouse over right corner square? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/97060-close-open-windows-when-mouse-over-right-corner-square/
; --------
; Minimize to Tray by Clicking Close Button? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/26521-minimize-to-tray-by-clicking-close-button/#entry176808
; ----
; ----
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
; ---- LOCATE ONLINE
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 41-Minimize Chrome Close & Close RButton.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2041-Minimize%20Chrome%20Close%20%26%20Close%20RButton.ahk
; ----
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; ----
; SOURCE
; ----
; Disable the minimize/maximize/close buttons - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/83455-disable-the-minimizemaximizeclose-buttons/
; ----
; -------------------------------------------------------------------
; FROM   Tue 16-Oct-2018 03:01:06
; TO     Tue 16-Oct-2018 03:30:00
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; ----
; SOURCE
; ----
; Close window when mouse hover red X square box? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/81233-close-window-when-mouse-hover-red-x-square-box/
; ----
; CODE TO MINIMIZE CHROME INSTEAD OF CLOSE 
; CODE ALSO RIGHT CLICK ON CLOSE WHICH NORMAL DOES NOTHING AND WHEN PRESS SHIFT KEY AND THEN WILL CLOSE
; CODE ADD _ TO USE WITH MULTIPLE PROGRAM 
; CODE ADD _ SOUNDBEEP WHENEVER ACTIVITY AT CLOSE BUTTON AND CLICK MADE
; CODE ADD _ THE TOP FEW LINE OF PIXEL IN CHROME GET TREAT AS CLOSE BUTTON 
; CODE ADD _ TOOLTIP INFO WORKER
; CODE ADD _ SPEED IT UP WITH CARE TO CPU USAGE FINALIZE
; -------------------------------------------------------------------
; FROM   Tue 16-Oct-2018 03:30:00
; TO     Tue 16-Oct-2018 08:54:00 __ 5 HOUR 24 MINUTE 
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 003
; -------------------------------------------------------------------
; ----
; AFTER THE ALL NIGHT SESSION FOR SIX HOUR AND ASLEEP IN 
; DAY UNTIL 3 PM DAY BEFORE AGAIN WE CODER
; FIXED IT SO WORKS ON WIN XP TASKBAR PROBLEM SOME RIGHT CLICK ITEM 
; LIKE VOLUME AND ONE OTHER NOT DISPLAY 
; LIKE SHOULD DO
; HAD GOOD LOOK AT CODE WHY THIS HAPPEN _ SORTED IT
; NOT SURE WHAT FIXER IN THE END BUT TIDIED UP _ PUT ENOUGH CODE IN
; ----
; CODE ADD _ THE WM_NCHITTEST := 0x0084 DOES NOT DETECT A CLOSE BUTTON UNDER CURSOR WHEN THE WINDOW IS 
; NOT IN FOCUS
; AND IS ABLE TO LEAD TO THE WRONG ACTION OF CLOSE WINDOW RATHER THAN SAFTEY MINIMIZE
; SO WE ON RIGHT OR LEFT CLICK BUTTON FORCE WINDOW INTO FOCUS 
; ----
; CODE ADD _ SORT THE PROBLEM WHERE HAD TO GET PARENT WINDOW WHEN DETECTING IF CHROME OR OTHER PROGRAM WANT INCLUDED
; ----
; CODE ADD _ MAKE THE TOOLTIP DISAPPEAR AFTER 2 SECOND HOVERING ON CLOSE BUTTON
; CODE ADD _ MAKE WORK FOR WIN XP WITHOUT AERO OFFSET SET AS MUCH
; -------------------------------------------------------------------
; FROM   Tue 16-Oct-2018 18:48:52
; TO     Tue 16-Oct-2018 20:22:00 __ 3 HOUR 34 MINUTE 
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 004
; -------------------------------------------------------------------
; ----
; DUE TO COMPILATION OF HEAVY CODE WHILE CLICKING MISTAKE WHILE 
; MACHINE UNDER HARD LOAD
; DECIDE NONE TOOLTIP ON OR SOUND EFFECT FOR WINDOWS THAT DON'T INVOLVE
; OKAY ALLOW SOUND EFFECT FOR ALL FROM HOVER WATCH_3 ROUTINE
;
; TIDY UP CODE JOB REFINE AND THING
; -------------------------------------------------------------------
; FROM   Wed 17-Oct-2018 09:44:06
; TO     Wed 17-Oct-2018 10:50:00 __ 1 HOUR 
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
; SESSION 005
; -------------------------------------------------------------------
; ----
; IMPROVEMENT WITH THE __ TRIGGER_HAPPEN=
; THE MOUSE WAS GETTING STICKY CLICK DOWN WHEN CLOSE WINDOW
;
; TIDY UP CODE JOB REFINE AND THING
; -------------------------------------------------------------------
; FROM   Wed 17-Oct-2018 23:26:46
; TO     Wed 17-Oct-2018 23:45:00
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 006
; -------------------------------------------------------------------
; ----
; DONE ALL THE FINE TUNE UP ALL IMPROVEMENT TO THE MAXIMUM FOR NOW
; TOOLTIPS COME BACK WHEN EVERY USE OF MINIMIZE AND CLOSE WINDOW IS DONE
; JUST TO LET KNOW STILL IN USE AND STILL THERE
; FINE TUNE CODE REMOVE REPRESENTATIVE CODE FINE TUNE REMOVED
; REDUNDANT USE OF SPAGHETTI CODE THAT WAS NEAR GENERATOR TIPPING POINT
; MUSH SMARTER
; IMPROVE THE SPEED I USE PROGRAMMER OF AUTOHOTKEYS
; BY HAVING A GOOD F12 KILL SWITCH FOR CODE THAT MESSING UP BEHAVING BADLY
; ONLY WHEN KEY HOOKING AND GENERALLY QUICKER WHEN RSI REPEAT ARRIVE
; PROCESS KILL BY PID NUMBER
; IMPROVE THE RESPONSE OF CODE
; BY TAKE PRIORITY THAT ACTIVE WINDOW RATHER THAN 
; HOVER OVER MOUSE COORDINATE
;
; WORKED LOT OF TIME BUT DONE 3 OR 4 PROJECT BETWEEN
; FOUND ONE LAST THING AS I WRITE TO CHECK OVER FOR MAXIMUM
; IN THE _ PROGRAM_SET_TO_USE _ ROUTINE - HARD TO DO THAT BIT OF EXIT EARLY
; WHEN FOUND RESULT COULD BE SUPERCEEDEDD CHANGED BY ANOTHER
; I COULD SAY PRIORTY IS TO THE ACTIVE WINDOW BUT DISABLE MOVE AROUND 
; A BIT HOOVER - REALLY I INTRO NEW CODE THAT COULD MAKE LESS QUICK AGAIN
; NOT DONE IS QUICKER OVER ACTIVE WINDOW
; DETECTION FOR FLOAT WINDOW AND ON CLOSE BUTTON IS THERE
;
; IMPROVED TOOLTIP THAT DISAPPEAR INSTANTLY WHEN WINDOW GONE WHILE 
; BEEN REST ON TIMER
; THATS THE MAXIMUM
;
; HAVE I GOT EVERYTHING I WANTED DONE FOR COMPUTER-ING TODAY

; SPEED THE DURATION OF SOUNDBEEP TO DOUBLE SPEEDIER NOT WHEN A MINIMIZE HAPPEN
; AS THAT LESS OFTEN
; MINIMIZE WAS SAFE FOR SOUND EVENT 
; BUT CLOSE NEEDED TAKE A CLOSER LOOK AT 
; AND PUT SOUND EVENT BACK ON AND SHORTED THEM BOTH 
; -------------------------------------------------------------------
; FROM   Sat 20-Oct-2018 00:45:42
; TO     Sat 20-Oct-2018 06:14:00 -----------------------------------
; FROM   Sat 20-Oct-2018 04:33:55 - MAIN THRASH OF CODE REALLY WAS
; TO     Sat 20-Oct-2018 06:24:00 - NEAR 2 HOUR
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
; SESSION 007
; -------------------------------------------------------------------
; ----
; LONGEST SLOG LEARN
; HAD TO GIVE UP CHASING PIXEL COLOUR LEARN AS HAVOC ON THE MOUSE CLICK EVENT
; REDONE ALL THE CODE MUCH MORE DIFFERENT AND BETTER
; THOUGHT OF EVERYTHING
; AND SPOTTED TON OF BUGGY FORM BEFORE
; AS ALTERNATIVELY I MANAGED TO SEE A WAY OF DOING WHAT WANTED BY LOOK AT TITLE
; YAHOO IN JUST DAYS GONE BY 
; CHANGED THE PRINT OUTPUT I USED TO GET HEADER INFO OF EMAIL
; TO INCLUDE TWO SCREEN WITH PRINT TO GO ALSO YOU ONLY HAVE TO PRESS ESCAPE ON THAT
; WITH SENDINPUT
; -------------------------------------------------------------------
; FROM   Mon 22-Oct-2018 20:32:40
; TO     Tue 23-Oct-2018 00:08:00
;# ------------------------------------------------------------------


;--------------------
#SingleInstance force
;--------------------
;--------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400

SETTIMER TIMER_PREVIOUS_INSTANCE,1

GLOBAL X
GLOBAL Y
GLOBAL HWND

GLOBAL O_X
GLOBAL O_Y

GLOBAL hWnd_APP

GLOBAL OSVER_N_VAR

GLOBAL TOOLTIP_DISPLAY_COUNT_LIMITER_1
GLOBAL TOOLTIP_DISPLAY_COUNT_LIMITER_2

GLOBAL CLOSE_BUTTON_HOVER_ACTIVITY_OLD

GLOBAL TRIGGER_FOR_WAIT_RIGHT_UP_HAPPEN
GLOBAL TRIGGER_FOR_WAIT_LEFT_UP_HAPPEN

TRIGGER_FOR_WAIT_RIGHT_UP_HAPPEN=FALSE
TRIGGER_FOR_WAIT_LEFT_UP_HAPPEN=FALSE

TOOLTIP_BEEN_SET_1=0
TOOLTIP_BEEN_SET_2=FALSE
TIMER_TOOLTIP := 0
TOOLTIP_DISPLAY_COUNT_LIMITER_1=0
TOOLTIP_DISPLAY_COUNT_LIMITER_2=0

HTCLOSE := 20
WM_NCHITTEST := 0x0084

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6
IF OSVER_N_VAR=10
	OSVER_N_VAR=10
	
SET_GO_2=0

O_X=0
O_Y=0


; SetTitleMatchMode 2  ; ANY
SetTitleMatchMode 3  ; EXACTLY


FN_Array := []
STAT_Array := []
WIN_CLASS_OR_TITLE_Array := []
ArrayCount := 0
  

; THIS ONE CAN TAKE OUT WHEN TESTER
; -------------------------------------------------------------------
; ArrayCount += 1
; FN_Array[ArrayCount]:="ThunderRT6FormDC"       ; ---- VISUAL BASIC
; STAT_Array[ArrayCount]:=FALSE
; WIN_CLASS_OR_TITLE_Array[ArrayCount]:="CLASS" 

ArrayCount += 1
FN_Array[ArrayCount]:="Chrome_WidgetWin_1"     ; ---- GOOGLE CHROME
STAT_Array[ArrayCount]:=TRUE
WIN_CLASS_OR_TITLE_Array[ArrayCount]:="CLASS"

; NOT WORKER ANYMORE YAHOO CHANGED PRINT MECHANISM
; MEAN ARE ABLE TO EXIT THE LOOP QUICKER WHEN CHECKER - MAYBE
; -------------------------------------------------------------------
; ArrayCount += 1
; FN_Array[ArrayCount]:="Print - Google Chrome"  ; ---- Print - Google Chrome
; STAT_Array[ArrayCount]:=FALSE
; WIN_CLASS_OR_TITLE_Array[ArrayCount]:="TITLE"

; GETTING PIXELS IS HARD WORK
; SO WORK AROUND DO THIS WAY ALLOW CLOSE WINDOW OF CHROME WHEN ON PRINTER OF BT YAHOO MAIL
; MUST FIND EXTRA SAFE GUARD
; -------------------------------------------------------------------
ArrayCount += 1
FN_Array[ArrayCount]:="BT Yahoo Mail -"          ; ---- Print - Google Chrome
STAT_Array[ArrayCount]:=FALSE
WIN_CLASS_OR_TITLE_Array[ArrayCount]:="TITLE"




; -------------------------------------------------------------------
; TOOLTIP INFO ABOUT CLOSE BUTTON HOVER
; AND EXTRA SOUNDBEEP-ER AT CLOSE ACTIVITY GETTING NEARER
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; BEGIN CODE WITH THIS TIMER
; BEGIN CODE ALSO HAPPEN WITH 
; MOUSE CLICK EVENT TWO MOUSE CLICK DOWN AND TWO MOUSE CLICK UP
; -------------------------------------------------------------------
SetTimer Watch_3, 100
; -------------------------------------------------------------------

Return

; -------------------------------------------------------------------
; END INIT
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; BEGIN CODE SUB ROUTINE-ER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
PROGRAM_SET_TO_USE:
; -------------------------------------------------------------------
SET_GO_2=FALSE
LOOP, 2
{	
	IF A_Index=1
	{
		WinGet, Hwnd_ID, ID, A
	}

	IF A_Index=2
	{
		CoordMode Mouse, Screen
		MouseGetPos, x, y, hWnd
		
		Hwnd_Parent := DllCall("GetParent", UInt,WinExist("ahk_id" hWnd)), Hwnd_Parent := !Hwnd_Parent ? WinExist("ahk_id" hWnd) : Hwnd_Parent
		Hwnd_ID:=Hwnd_Parent
	}
	
	WinGetClass, Class_Title, ahk_id %Hwnd_ID%
	WinGetTitle, Win_Title, ahk_id %Hwnd_ID%

	
	Loop % ArrayCount
	{
		; BIT OF CROSS CHECKER HERE SPEDIER WITH AN ELSE LINE
		; -----------------------------------------------------------
		IF WIN_CLASS_OR_TITLE_Array[A_Index]="CLASS"
			IF FN_Array[A_Index]=Class_Title
				SET_GO_2:=% STAT_Array[A_Index]
				
		IF WIN_CLASS_OR_TITLE_Array[A_Index]="TITLE"
			IF (InStr("*"Win_Title,"*"FN_Array[A_Index]))
				SET_GO_2:=% STAT_Array[A_Index]
	}

	IF SET_GO_2=1
		SET_GO_2=TRUE
	
	IF SET_GO_2=0
		SET_GO_2=FALSE

	; ---------------------------------------------------------------
	; TEMPORARY ABORTED USE THIS METHOD
	; TOO DIFFICULT FOR CLICK AND DRAGGING
	; ---------------------------------------------------------------
	IF A_Index=4
		IF Class_Title=Chrome_WidgetWin_1
		{
			X=180
			Y=55
			PixelGetColor Color_1, %X%, %Y%, RGB
			; X=200
			; PixelGetColor Color_2, %X%, %Y%, RGB
			; X=220
			; PixelGetColor Color_3, %X%, %Y%, RGB
			X=24
			Y=18
			PixelGetColor Color_4, %X%, %Y%, RGB
			
			; ---------------------------------------------------------------
			; THIS THE SELECTED COLOR IN THE TITLE URL BAR
			; BUT NEVER HAPPEN
			; MUST BE WHEN UNSELECTED BIT MORE GREY THAN WHITE
			; ALSO CHANGES COLOR GREY-ER A BIT WHEN HOVER OVER
			; ---------------------------------------------------------------
			; FIND CHECKING TOO MANY POINT MAKE PROBLEM HAZARD TO COPY SELECT PICKER
			; REDUCE BY TWO AND OKAY
			; ALSO INCLUDE ONLY CHECK ON ACTIVE WINDOW SPEEDIER
			; ---------------------------------------------------------------
			COLOR_COMPARE_1=0xF1F3F40xDE736A
			COLOR_COMPARE_2=%COLOR_1%%COLOR_4%
			IF COLOR_COMPARE_1=%COLOR_COMPARE_2%
			{
				SET_GO_2=FALSE
				RETURN
			}

			
		}
		; TOOLTIP % SET_GO_2 " -- " COLOR_COMPARE_1 " -- " COLOR_COMPARE_2
		; TOOLTIP % SET_GO_2
}


	
RETURN
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; -------------------------------------------------------------------
LButton:: ; Minimize Google Chrome instead of close when close button is clicked
; -------------------------------------------------------------------

	GOSUB PROGRAM_SET_TO_USE
		
	IF SET_GO_2=FALSE 
	{
		Click down
		RETURN
	}
	
	; --------------------------------------------------
	; NICE TRY TO SET FOCUS TO WINDOW UNDER MOUSE CURSOR 
	; TOO MANY PROBLEM LIKE LEFT CLICK ON TASKBAR
	; FOUND THE PROBLEM USE THE NEW ROUTINE Hwnd_Parent IN PROGRAM_SET_TO_USE
	; --------------------------------------------------
	; CoordMode Mouse, Screen
	; MouseGetPos, x, y, hWnd
	; Hwnd_Parent := DllCall("GetParent", UInt,WinExist("ahk_id" hWnd)), Hwnd_Parent := !Hwnd_Parent ? WinExist("ahk_id" hWnd) : Hwnd_Parent
	; WinGetClass, ClassNN_1, ahk_id %hwnd_parent%
	; WinGetClass, ClassNN_2, A
	; WinGet, Hwnd_2, ID, A
	; TOOLTIP % ClassNN_1 " -- " ClassNN_2
	; TOOLTIP % SET_GO_2 " -- " ClassNN_1 " -- " ClassNN_2


	GOSUB IsOverCloseButton

	; -----------------------------------------------------------
	; PUT TOOLTIP BACK ON IF USED THE FUNCTION REMINDER STILL THERE
	; -----------------------------------------------------------
	IF SET_GO_1=TRUE
	{
		TOOLTIP_DISPLAY_COUNT_LIMITER_1=0
		TOOLTIP_DISPLAY_COUNT_LIMITER_2=0
	}
		
	TRIGGER_HAPPEN=FALSE
	IF SET_GO_2=TRUE
		If SET_GO_1=TRUE
		{
			; SOUNDBEEP 5000,100
			; -----------------------------------------------------------
			; MMX 0 = NORMAL -- MMX 1 = MAXIMIZED -- MMX -1 = MINIMIZED
			; -----------------------------------------------------------
			WinGet MMX, MinMax, ahk_id %hWnd_APP%
			If MMX>-1
			{
				WinMinimize ahk_id %hWnd_APP%
				SOUNDBEEP 2000,100
				TRIGGER_HAPPEN=TRUE
				TRIGGER_FOR_WAIT_LEFT_UP_HAPPEN=TRUE
			}
		}

	IF TRIGGER_HAPPEN=FALSE		
		Click down

RETURN
; -------------------------------------------------------------------

LButton Up::
; -------------------------------------------------------------------
IF TRIGGER_FOR_WAIT_LEFT_UP_HAPPEN=TRUE
{
	TRIGGER_FOR_WAIT_LEFT_UP_HAPPEN=FALSE
	RETURN
}
Click Up
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
RIGHT_CLICK_CLOSE_IT:
	; -------------------------------------------------------------------
	GOSUB PROGRAM_SET_TO_USE
	IF SET_GO_2=FALSE 
	{
		Click, down, right
		RETURN
	}

	GOSUB IsOverCloseButton

	; -----------------------------------------------------------
	; PUT TOOLTIP BACK ON IF USED THE FUNCTION REMINDER STILL THERE
	; -----------------------------------------------------------
	IF SET_GO_1=TRUE
	{
		TOOLTIP_DISPLAY_COUNT_LIMITER_1=0
		TOOLTIP_DISPLAY_COUNT_LIMITER_2=0
	}
		
	TRIGGER_HAPPEN=FALSE
	IF SET_GO_2=TRUE
		If SET_GO_1=TRUE
		{
			; SLEEP 400 
			; -----------------------------------------------------------
			; IF NOT A DELAY HERE THEN RIGHT CLICK WILL HAPPEN ON 
			; THE WINDOW UNDERNEATH THE ONE THAT CLOSED
			; -----------------------------------------------------------
			; BETTER TO USE KeyWait, RButton 
			; BUT HAVE TO DO BEFORE WINCLOSE
			; AND SLIGHT MISLEADING EXCPT FOR BEEP THAT NOTHING GOING TO HAPPEN 
			; AS EVENT IS AFTER RBUTTON UP
			; BUT SOUNDBEEP OUGHT TO GIVE A CLUE
			; -----------------------------------------------------------
			; KeyWait, RButton
			WinClose ahk_id %hWnd_APP%
			TRIGGER_HAPPEN=TRUE
			TRIGGER_FOR_WAIT_RIGHT_UP_HAPPEN=TRUE
			SOUNDBEEP 2000,40

		}

	IF TRIGGER_HAPPEN=FALSE		
		Click, down, right

RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
RButton Up::
; -------------------------------------------------------------------
IF TRIGGER_FOR_WAIT_RIGHT_UP_HAPPEN=TRUE
{
	; ---------------------------------------------------------------
	; STOP PRESS BUTTON ON WINDOW UNDERNETH WHEN CLOSE HAS BEEN MADE TO HAPPEN
	; BEFORE I WAS USE KEYWAIT UP BUT THINK THAT MADE TOO MANY PROBLEM
	; HOLDING UP
	; ---------------------------------------------------------------
	TRIGGER_FOR_WAIT_RIGHT_UP_HAPPEN=FALSE
	RETURN
}
Click, up, right
RETURN



; -------------------------------------------------------------------

; CODE WAS IN PROCEDURE ABOVE BUT REDUNDANT NOW
; -------------------------------------------------------------------
IF GetKeyState("RButton")=TRUE
{
	Click, up, right
	RETURN
}
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
Watch_3:
; -------------------------------------------------------------------
; REFINE THE CPU USAGE AS FINAL PART DONE
; -------------------------------------------------------------------
IF TIMER_TOOLTIP>0
	IF TIMER_TOOLTIP < %A_NOW%
	{
		TOOLTIP
		TIMER_TOOLTIP:=0
		TOOLTIP_BEEN_SET_1=0
	}

GOSUB PROGRAM_SET_TO_USE

IF (O_X=%X% and O_Y=%Y%)
	Return
O_X=X
O_Y=Y

; ----------------------------------------------------------------
; IF WANT SINGLE ACTION HOVER ON ONLY DON'T REQUIRE THIS LINE CODE
; AND ALSO DO WANT TO INCLUDE NEXT IF LINE
; AND VICE VERSA __ AKE ONE OUT AND PUT ONE BACK
; IF WANT DOUBLE ACTION BEEP ON HOVER OFF AND ON DO THE SWAP
; DON'T LEAVE A LINE BREAK BETWEEN THE TWO IF LINE - JOINER
; AND DO YOUR INDENTING CODER - IT SAVES CODE
; ----------------------------------------------------------------

GOSUB IsOverCloseButton
CLOSE_BUTTON_HOVER_ACTIVITY=%SET_GO_1%

IF CLOSE_BUTTON_HOVER_ACTIVITY=FALSE
	TIMER_TOOLTIP = 1
	
; MSGBOX % CLOSE_BUTTON_HOVER_ACTIVITY

IF CLOSE_BUTTON_HOVER_ACTIVITY_OLD<>%CLOSE_BUTTON_HOVER_ACTIVITY%
IF CLOSE_BUTTON_HOVER_ACTIVITY=TRUE
{
	SOUNDBEEP 1000,20
}
CLOSE_BUTTON_HOVER_ACTIVITY_OLD=%CLOSE_BUTTON_HOVER_ACTIVITY%
; ---------------------------------------------------------------
; ---------------------------------------------------------------
	
IF CLOSE_BUTTON_HOVER_ACTIVITY=TRUE
{
	IF SET_GO_2=TRUE
	{
		IF TOOLTIP_BEEN_SET_1=0
		{
			TOOLTIP_DISPLAY_COUNT_LIMITER_1+=1
			IF TOOLTIP_DISPLAY_COUNT_LIMITER_1>5
				RETURN
			ToolTip % "LEFT MOUSE BUTTON = MINIMIZE`r`nRIGHT MOUSE BUTTON = CLOSE"
			TOOLTIP_BEEN_SET_1=1
			TOOLTIP_BEEN_SET_2=TRUE
			TIMER_TOOLTIP = % A_Now
			TIMER_TOOLTIP += 4, seconds
		}
	}

	IF SET_GO_2=FALSE
	{
		; -------------------------------------------------------------------
		; THIS SECOND TOOLTIP FOR ANY WINDOW JUST TO SHOW HOVER 
		; NOT GETTING USE ANYMORE ONLY BEEP WAS REQUIRING
		; PULL OUT FROM JUST ABOVE -- IF SET_GO_2=FALSE 
		; -------------------------------------------------------------------
		IF TOOLTIP_BEEN_SET_1=0
		{
			TOOLTIP_DISPLAY_COUNT_LIMITER_2+=1
			IF TOOLTIP_DISPLAY_COUNT_LIMITER_2>5
				RETURN
			ToolTip % "MOUSE`r`nON`r`nCLOSE`r`nBUTTON"
			TOOLTIP_BEEN_SET_1=2
			TOOLTIP_BEEN_SET_2=TRUE
			TIMER_TOOLTIP = % A_Now
			TIMER_TOOLTIP += 4, seconds
		}
	}
}
Return
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; CORRECT METHOD FOUND FOR AREO
; ----
; Minimize to Tray by Clicking Close Button? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/26521-minimize-to-tray-by-clicking-close-button/#entry176808
; ----

; Reference: http://www.autohotkey.com/board/topic/20431-wm-nchittest-wrapping-whats-under-a-screen-point/
; -------------------------------------------------------------------
IsOverCloseButton:
; -------------------------------------------------------------------
{
	
	MouseGetPos,,,WIN_ID_UNDER_MOUSE_CURSOR
	hWnd_APP=%WIN_ID_UNDER_MOUSE_CURSOR%
	WinGetClass, Class_Title, ahk_id %WIN_ID_UNDER_MOUSE_CURSOR%
	IfInString, Class_Title, Shell_TrayWnd
	{
		SET_GO_1=FALSE
		Return
	}
	
	; ---------------------------------------------------------------
	; THE HIGHER THE X OFFSET NUMBER THE MORE CLOSER IT IS TO THE 
	; CLOSE BUTTON AND NOT SPREAD ALSO ONTO MAXIMIZE
	; ---------------------------------------------------------------
	; DON'T USE THE SAME OFFSET WHEN IN WINDOWS XP
	; ---------------------------------------------------------------
	IF OSVER_N_VAR > 5 ; ---- ABOVE WIN XP 
		x-=30

	IF OSVER_N_VAR = 5 ; ---- WIN XP 
		x-=19
	
	; ---------------------------------------------------------------
	; CHROME HAS GOT MORE OVERLAP ONTO MAXIMIZE BUTTON
	; ANOTHER OFFSET
	; CAN'T BLOCK THE MAXIMIZE BUTTON OVER
	; ---------------------------------------------------------------
	WinGetClass, Class_Title, ahk_id %hWnd%
	IfInString, Class_Title, Chrome_WidgetWin_1
		IF OSVER_N_VAR > 5 
			x-=25
		
	SET_EXIT_VAR=FALSE
	; WM_NCHITTEST := 0x0084
	SendMessage, 0x84,, (x & 0xFFFF) | (y & 0xFFFF) << 16,, ahk_id %hWnd%
	If ErrorLevel in 9,12,14,20 ; 20 is for Close box but Vista/Aero bug in WM_NCHITTEST	
	{
		SET_EXIT_VAR=TRUE
	}
	
	CoordMode Mouse, Window
	MouseGetPos, x, y, hWnd
	WinGetPos, WinPosX, WinPosY, WindowWidth, WindowHeight, ahk_id %hWnd%
	
	; ---------------------------------------------------------------
	; SOME EXTRA
	; SAFEGUARD THE TOP FEW LINE OF PIXEL IN CHROME GET TREAT AS CLOSE BUTTON 
	; ACCORDING TO WM_NCHITTEST
	; SO WE HAVE NAILED IT AGAIN
	; ---------------------------------------------------------------
	IF (y>45 or WindowWidth-x>80)
		SET_EXIT_VAR=FALSE

	; FALSE NOT A WINDOW WANT TO TAKE ACTION ON
	; ---------------------------------------------------------------
	IF SET_EXIT_VAR=FALSE
		SET_GO_1=FALSE

		
	; TOOLTIP % WindowWidth " -- " WindowHeight  " -- " x " -- " y " -- " WindowWidth-x " -- " SET_EXIT_VAR
		
	IF SET_EXIT_VAR=TRUE
		SET_GO_1=TRUE
	ELSE
		SET_GO_1=FALSE
	
}
RETURN
; -------------------------------------------------------------------




; ---------------------------------------------------------------------
; THIS ROUTINE COULD BE INCLUDED IN HERE BUT IS NOT CALLED FROM ANYWHERE
; RUN FORM ANOTHER CODE ON TIMER
; Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
; ---------------------------------------------------------------------
SUB_TO_CLOSE_PRINT_WINDOW:
; -------------------------------------------------------------------
; THE NEW YAHOO MAIL INTERFACE HAS PROBLEM
; WHEN I WANT HEADER INFO OF AN EMAIL LIKE IT WAS TO PRINT ONE
; TO COPY PASTE THE TEXT HEADER
; NOW THE NEW PRINT OF YAHOO INCLUDE DOBLE CALL WINDOW TO PRINT ALSO
; WHEN THAT WAS OPTION ON NICER FIRST SCREEN WINDOW BEFORE
; SO I MAKE A PIXEL FIND COLOR AND CLOSE PRINT SCREEN 
; SAVE ME HAVE TO
; [ Monday 20:08:30 Pm_22 October 2018 ]
; ANOTHER LITTLE PROBLEM SOLVED
; READY FOR SOME OTHER THING IN CHROME WORKER I LIKING
; [ Monday 20:08:30 Pm_22 October 2018 ]
; -------------------------------------------------------------------

FN_Array_1 := []
FN_Array_2 := []

WinGet, id, list,ahk_class Chrome_WidgetWin_1
Loop, %id%
{
	table := id%A_Index%
	WinGetTitle, title, ahk_id %table%
	FN_Array_1[A_Index] := "%table%"
	FN_Array_2[A_Index] := id%A_Index%
}

CoordMode Pixel, Window 

Loop % FN_Array_1.MaxIndex()
	ArrayCount_VAR_1=%A_Index%
	Loop % FN_Array_1.MaxIndex()
	{
		ArrayCount_VAR_2=%A_Index%
		IF ArrayCount_VAR_1<>ArrayCount_VAR_2    ; NOT THE SELF
		{
			IF FN_Array_1[ArrayCount_VAR_1]=FN_Array_1[ArrayCount_VAR_2]
				;MSGBOX "HOO"
				; MSGBOX % FN_Array_2[ArrayCount_VAR_2]
				HWND_4 = % FN_Array_2[ArrayCount_VAR_2]
				WinGet, HWND_2, ID, ahk_id %HWND_4%
				IF HWND_4>0
				{
					X=500
					Y=100
					PixelGetColor Color_1, %X%, %Y%, RGB
					X=242
					Y=180
					PixelGetColor Color_2, %X%, %Y%, RGB
					X=295
					Y=180
					PixelGetColor Color_3, %X%, %Y%, RGB
					
					COLOR_COMPARE_1=0x5256590x4B8EF90x4B8EF9
					COLOR_COMPARE_2=%COLOR_1%%COLOR_2%%COLOR_3%
					IF COLOR_COMPARE_1=%COLOR_COMPARE_2%
					{
						SOUNDBEEP 1000,50
						; TOOLTIP % COLOR_COMPARE_1 " -- " COLOR_COMPARE_2
						; WinCLOSE  ahk_id %HWND_4%
						SENDINPUT {ESCAPE}
						RETURN
					}
						; WinMinimize  ahk_id %HWND_4%
			}
		}
	}
; ----
; PixelGetColor from an innactive window - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/103880-pixelgetcolor-from-an-innactive-window/
; ----

RETURN



; -------------------------------------------------------------------
; OLDER CODE WHEN GOT STATING
; -------------------------------------------------------------------
; girlgamer
; Moderators
; 3263 posts
; Last active: Feb 01 2015 09:49 AM
; Joined: 04 Jun 2010
; ----
; Disable the minimize/maximize/close buttons - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/83455-disable-the-minimizemaximizeclose-buttons/
; ----
; -------------------------------------------------------------------
; F12:: ;<-- Move win and change style
	; WinGet, TempWindowID, ID, A
	; If (WindowID != TempWindowID)
	; {
		; WindowID:=TempWindowID
		; WindowState:=0
	; }
	; If (WindowState != 1)
	; {
		; WinGetPos, WinPosX, WinPosY, WindowWidth, WindowHeight, ahk_id %WindowID%
		; WinSet, Style, ^0xC40000, ahk_id %WindowID%
		; ;WinMove, ahk_id %WindowID%, , 0, 0, A_ScreenWidth, A_ScreenHeight
	; }
	; Else
	; {
		; WinSet, Style, ^0xC40000, ahk_id %WindowID%
		; ;WinMove, ahk_id %WindowID%, , WinPosX, WinPosY, WindowWidth, WindowHeight
	; }
	; WindowState:=!WindowState
; return
; -------------------------------------------------------------------
  
; -------------------------------------------------------------------
; TAKEN OUT NOT GOOD IDEA TO USE TWO OPTION OF RIGHT CLICK
; -------------------------------------------------------------------
; EITHER USE RBUTTON ON ITS OWN OR AS ABOVE SHIFT KEY AS SHOW BUT REQUIRE KEYBOARD AROUND
; -------------------------------------------------------------------
; ~Shift & RButton:: ; Minimize Google Chrome instead of close when close button is clicked
; ; -------------------------------------------------------------------
; GOSUB RIGHT_CLICK_CLOSE_IT
; return
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; EITHER USE RBUTTON ON ITS OWN OR AS ABOVE SHIFT KEY ALSO BUT REQUIRE KEYBOARD AROUND
; -------------------------------------------------------------------
~RButton:: ; Minimize Google Chrome instead of close when close button is clicked
; -------------------------------------------------------------------
	GOSUB RIGHT_CLICK_CLOSE_IT

	RETURN
; -------------------------------------------------------------------

; CODE WAS IN PROCEDURE RButton Up:: ABOVE BUT REDUNDANT NOW
; -------------------------------------------------------------------
; IF GetKeyState("LButton")=TRUE
; {
	; Click Up
	; RETURN
; }
; RETURN

  
 
 MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		DetectHiddenWindows, On 
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List% 
		  { 
			WinGet, PID, PID, % "ahk_id " List%A_Index% 
			If ( PID <> DllCall("GetCurrentProcessId") ) 
				 ; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				 Process, Close, List%A_Index% 
		  }
		Process, Close,% DllCall("GetCurrentProcessId")
		
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
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return

 
 
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
; REFERENCE PAGES OPEN 30
; -------------------------------------------------------------------
; ----
; AHK DISABLE CHROME CLOSE BUTTON - Google Search
; https://www.google.co.uk/search?q=AHK+DISBALE+CHROME+CLOSE+BUTTON&oq=AHK+DISBALE+CHROME+CLOSE+BUTTON&aqs=chrome..69i57j69i64.7256j0j7&sourceid=chrome&ie=UTF-8
; --------
; Disable the minimize/maximize/close buttons - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/83455-disable-the-minimizemaximizeclose-buttons/
; --------
; AHK MOVE MOUSE WHEN OVER CLOSE BUTTON - Google Search
; https://www.google.co.uk/search?q=AHK+MOVE+MOUSE+WHEN+OVER+LCOSE+BUTTON&oq=AHK+MOVE+MOUSE+WHEN+OVER+LCOSE+BUTTON&aqs=chrome..69i57j69i64.7656j0j7&sourceid=chrome&ie=UTF-8
; --------
; Close window when mouse hover red X square box? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/81233-close-window-when-mouse-hover-red-x-square-box/
; --------
; Close open windows when mouse over right corner square? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/97060-close-open-windows-when-mouse-over-right-corner-square/
; --------
; detect button under cursor (using WM_NCHITTEST)? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/57049-detect-button-under-cursor-using-wm-nchittest/
; --------
; Minimize to Tray by Clicking Close Button? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/26521-minimize-to-tray-by-clicking-close-button/#entry176808
; --------
; Disable And Hide Close Button On Gui - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/59515-disable-and-hide-close-button-on-gui/
; --------
; disable "window close button (X)" - Google Groups
; https://groups.google.com/forum/#!topic/vim_use/lh0JTeQ5Tws
; --------
; Downloads – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/downloads/
; --------
; CloseToQuit – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/closetoquit/
; --------
; DelEmpty – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/delempty/
; --------
; BatteryRun – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/batteryrun/
; --------
; CacheSort – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/cachesort/
; --------
; NoClose – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/noclose/
; --------
; NoClose – Skrommel's One Hour Software
; http://www.dcmembers.com/skrommel/download/noclose/
; --------
; WinSet - Syntax & Usage | AutoHotkey
; https://www.autohotkey.com/docs/commands/WinSet.htm
; --------
; MouseMove - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/MouseMove.htm
; --------
; Loop - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/Loop.htm
; --------
; Nested Loops - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/62226-nested-loops/
; --------
; Adding Variables - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/73484-adding-variables/
; --------
; Want to use RButton as a shift key for keyboard - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/57907-want-to-use-rbutton-as-a-shift-key-for-keyboard/
; --------
; Click - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/Click.htm
; --------
; autohotkey - Do something while shift and left click pressed or right click pressed - Stack Overflow
; https://stackoverflow.com/questions/32438778/do-something-while-shift-and-left-click-pressed-or-right-click-pressed?rq=1
; --------
; WinClose - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/WinClose.htm
; --------
; Sense if RButton being held down (for X amount of time?) - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/100802-sense-if-rbutton-being-held-down-for-x-amount-of-time/
; --------
; KeyWait - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/KeyWait.htm
; --------
; Disabled right click. Can't abort script. - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/123104-disabled-right-click-cant-abort-script/
; --------
; CoordMode - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/CoordMode.htm
; --------
; WinGetPos - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/WinGetPos.htm
; ----
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 003
; REFERENCE PAGES OPEN 09
; -------------------------------------------------------------------
; ----
; how to find the window title of the window under mouse? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/9380-how-to-find-the-window-title-of-the-window-under-mouse/
; --------
; GetKeyState() / GetKeyState - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/GetKeyState.htm
; --------
; FileSetTime - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/FileSetTime.htm#YYYYMMDD
; --------
; add time to %a_now% - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/79835-add-time-to-a-now/
; --------
; WinGetClass - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/WinGetClass.htm
; --------
; Get parent window posssible? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/65666-get-parent-window-posssible/
; --------
; Get the Parent HwnD when selecting a i.e. a ComboBox - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/57256-get-the-parent-hwnd-when-selecting-a-ie-a-combobox/
; --------
; What is ahk_id? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/50667-what-is-ahk-id/
; --------
; WinGet - Syntax & Usage | AutoHotkey
; https://autohotkey.com/docs/commands/WinGet.htm#SubCommands
; ----
; -------------------------------------------------------------------
