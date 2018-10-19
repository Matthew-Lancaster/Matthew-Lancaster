;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk
;# __ 
;# __ Autokey -- 41-Minimize Chrome Close & Close RButton.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Tue 16-Oct-2018 03:00:00
;# __ 
;  =============================================================

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
; CODE ADD _ MAKE THE TOOLTIP DISAPPEAR AFTER 2 SECOND HOOVERING ON CLOSE BUTTON
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
; OKAY ALLOW SOUND EFFECT FOR ALL FROM HOOVER WATCH_3 ROUTINE
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
; APROVEMENT WITH THE __ TRIGGER_HAPPEN=
; THE MOUSE WAS GETTING STICKY CLICK DOWN WHEN CLOSE WINDOW
;
; TIDY UP CODE JOB REFINE AND THING
; -------------------------------------------------------------------
; FROM   Wed 17-Oct-2018 23:26:46
; TO     Wed 17-Oct-2018 23:45:00
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

GLOBAL CLOSE_BUTTON_HOOVER_ACTIVITY_OLD


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

; -------------------------------------------------------------------
; TOOLTIP INFO ABOUT CLOSE BUTTON HOOVER
; AND EXTRA SOUNDBEEP-ER AT CLOSE ACTIVITY GETTING NEARER
; -------------------------------------------------------------------
SetTimer Watch_3, 100
; -------------------------------------------------------------------

Return

; -------------------------------------------------------------------
; END INIT
; -------------------------------------------------------------------
; BEGIN CODE
; -------------------------------------------------------------------


; -------------------------------------------------------------------
PROGRAM_SET_TO_USE:
; -------------------------------------------------------------------
	
	CoordMode Mouse, Screen
	MouseGetPos, x, y, hWnd
	
	Hwnd_Parent := DllCall("GetParent", UInt,WinExist("ahk_id" hWnd)), Hwnd_Parent := !Hwnd_Parent ? WinExist("ahk_id" hWnd) : Hwnd_Parent

	WinGetClass, Class_Title, ahk_id %Hwnd_Parent%
	WinGetTitle, Win_Title, ahk_id %Hwnd_Parent%
	
	; BE MORE PRECISE
	SET_GO_2=FALSE
	; VISUAL BASIC
	If Class_Title=ThunderRT6FormDC
		SET_GO_2=false
	
	; GOOGLE CHROME
	If Class_Title=Chrome_WidgetWin_1
		SET_GO_2=TRUE
	
	If Win_Title=Print - Google Chrome
		SET_GO_2=FALSE
	
	; ; VISUAL BASIC
	; IfInString, Class_Title, ThunderRT6FormDC


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

SET_GO_1=FALSE
IF IsOverCloseButton(X, Y, hWnd)
	SET_GO_1=TRUE

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

TRIGGER_HAPPEN=FALSE
IF SET_GO_2=TRUE
	If SET_GO_1=TRUE
	{
		WinMinimize ahk_id %hWnd_APP%
		SOUNDBEEP 2000,100
		TRIGGER_HAPPEN=TRUE
	}

; THIS PART NEVER HAPPEN ANYMORE AS ABOVE
; -------------------------------------------------------------------
IF SET_GO_1=TRUE
	IF SET_GO_2=FALSE
		SOUNDBEEP 1000,100

IF TRIGGER_HAPPEN=FALSE		
	Click down
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

SET_GO_1=FALSE
if IsOverCloseButton(X, Y, hWnd)
	SET_GO_1=TRUE

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
		; SOUNDBEEP 2000,100
		KeyWait, RButton
		WinClose ahk_id %hWnd_APP%
		TRIGGER_HAPPEN=TRUE
	}

; THIS PART NEVER HAPPEN ANYMORE AS ABOVE
; -------------------------------------------------------------------
IF SET_GO_1=TRUE
	IF SET_GO_2=FALSE
		; SOUNDBEEP 1000,100

IF TRIGGER_HAPPEN=FALSE		
Click, down, right

RETURN
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
; IF GetKeyState("Shift")=TRUE
	; Return
GOSUB RIGHT_CLICK_CLOSE_IT
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
LButton Up::
; -------------------------------------------------------------------
Click Up
RETURN
; -------------------------------------------------------------------

; CODE WAS IN PROCEDURE ABOVE BUT REDUNDANT NOW
; -------------------------------------------------------------------
IF GetKeyState("LButton")=TRUE
{
	Click Up
	RETURN
}
RETURN

; -------------------------------------------------------------------

; -------------------------------------------------------------------
RButton Up::
; -------------------------------------------------------------------
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
	}

GOSUB PROGRAM_SET_TO_USE
IF SET_GO_2=FALSE 
{
	; CLOSE_BUTTON_HOOVER_ACTIVITY_OLD=FALSE
	; RETURN
}

IF (O_X=%X% and O_Y=%Y%)
	Return
O_X=X
O_Y=Y

CLOSE_BUTTON_HOOVER_ACTIVITY=FALSE
if IsOverCloseButton(X, Y, hWnd)
	CLOSE_BUTTON_HOOVER_ACTIVITY=TRUE

	IF CLOSE_BUTTON_HOOVER_ACTIVITY_OLD=%CLOSE_BUTTON_HOOVER_ACTIVITY%
		RETURN

	IF CLOSE_BUTTON_HOOVER_ACTIVITY_OLD<>%CLOSE_BUTTON_HOOVER_ACTIVITY%
	IF CLOSE_BUTTON_HOOVER_ACTIVITY=TRUE
	{
		SOUNDBEEP 1000,40
	}
		
	CLOSE_BUTTON_HOOVER_ACTIVITY_OLD=%CLOSE_BUTTON_HOOVER_ACTIVITY%

	IF SET_GO_2=FALSE 
		RETURN

	IF CLOSE_BUTTON_HOOVER_ACTIVITY=TRUE
	{
		IF SET_GO_2=TRUE
		{
			IF TOOLTIP_BEEN_SET_1<>1
			{
				TOOLTIP_DISPLAY_COUNT_LIMITER_1+=1
				IF TOOLTIP_DISPLAY_COUNT_LIMITER_1>4
					RETURN
				ToolTip % "LEFT MOUSE BUTTON = MINIMIZE`r`nRIGHT MOUSE BUTTON = CLOSE"
				TOOLTIP_BEEN_SET_1=1
				TOOLTIP_BEEN_SET_2=TRUE
				TIMER_TOOLTIP := a_now + 2
			}
		}
		ELSE
		{
			IF TOOLTIP_BEEN_SET_1<>2
			{
				TOOLTIP_DISPLAY_COUNT_LIMITER_2+=1
				IF TOOLTIP_DISPLAY_COUNT_LIMITER_2>10
					RETURN
				ToolTip % "MOUSE`r`nON`r`nCLOSE`r`nBUTTON"
				TOOLTIP_BEEN_SET_1=2
				TOOLTIP_BEEN_SET_2=TRUE
				TIMER_TOOLTIP := a_now + 2
			}
		}
	}
	ELSE
	{
		IF TOOLTIP_BEEN_SET_2=TRUE
		{
			TOOLTIP
			TOOLTIP_BEEN_SET_1=0
			TOOLTIP_BEEN_SET_2=FALSE
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
IsOverCloseButton(x, y, hWnd)
; -------------------------------------------------------------------
{
	hWnd_APP=%hWnd%
	
	MouseGetPos,,,WIN_ID_UNDER_MOUSE_CURSOR
	WinGetClass, Class_Title, ahk_id %WIN_ID_UNDER_MOUSE_CURSOR%
	IfInString, Class_Title, Shell_TrayWnd
	{
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
		SET_EXIT_VAR=TRUE
	
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
	
	; TOOLTIP % WindowWidth " -- " WindowHeight  " -- " x " -- " y " -- " WindowWidth-x " -- " SET_EXIT_VAR

	IF SET_EXIT_VAR=TRUE
		RETURN TRUE
}
RETURN
; -------------------------------------------------------------------


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
