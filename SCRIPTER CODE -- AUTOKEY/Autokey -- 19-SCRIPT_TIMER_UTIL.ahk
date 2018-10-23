;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
;# __ 
;# __ Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Tue 00:52:00 Am_12 Dec 2017 ]
;# LAST EDIT TIME [ Fri 16:08:00 Pm_27 Apr 2018 ]
;# __ 
;  =============================================================


; -------------------------------------------------------------------
; 002
; -------------------------------------------------------------------
; FROM TO Sun 15-Apr-2018 19:26:10
; -------------------------------------------------------------------
; THIS IS OUR PREFERRED DEFAULT OPTIONS FOR INSTALLING NOTEPAD++
; THE MIDDLE CHECKBOX IS SELECTED
; Reference's at End
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 003
; -------------------------------------------------------------------
; New work the GoodSync Routine to Set Number of Hour in Options
; From Default 2 to 4
; Straight Form The Help File Nothing Checked Online
; -------------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 13:44:49
; TO ---- Fri 27-Apr-2018 14:28:00 -- 50 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Well That Hard Work To Set the Content of an Edit Control 
; You Must Delete Whats in There First to Do with Caret Positioning
; And I Had to Look Up on the Internet
; Better than Using Sendinput key and Having to Be In Focus 
; Active Window __ Very Easy in the End
; Credit Due
; ----
; Delete edit control text ahk - Google Search
; https://www.google.co.uk/search?num=50&rlz=1C1CHWL_enGB794&ei=HDjjWvX4KI63gQab46zQDQ&q=Delete+edit+control+text+ahk&oq=Delete+edit+control+text+ahk&gs_l=psy-ab.3..35i39k1.5379.5647.0.6669.2.2.0.0.0.0.143.203.1j1.2.0....0...1c.1.64.psy-ab..0.2.201....0.9781WNbiPGQ
; --------
; Delete contents of edit control? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/80399-delete-contents-of-edit-control/
; ----
; -------------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 14:28:00
; TO ---- Fri 27-Apr-2018 15:51:00 -- 1 HOUR 20 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 004
; MORE WORK WANTED A SOUND EFFECT IF A WINDOW CAME OUT OF NOT RESPONDING 
; TO LET LEARN THE WAIT IS OVER FOR GOODSYNC NETWORK PATHS TIME CONSUMING
; TOOK A WHILE BECAUSE SOUND SYSTEM EVENTS WERE SWITCHED OFF
; -------------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 18:53:27
; TO ---- Fri 27-Apr-2018 19:33:58 -- 40 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 005
; NEXT WANTER THE CHECKBOX SET WHEN GOODSYNC HAS THE TEXT BOX FILLED
; AND IT CHECK HWND OF WINDOW AND DOES ONCE 
; IT HARD TO TEST BECAUSE WHEN CHECKBOX NOT CHECKED THE TEXTBOX FIGUAR 
; ENTRY IS ENABLED=FALSE
; -------------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 19:33:58
; TO ---- Fri 27-Apr-2018 19:53:05 -- 20 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 00*
; SESSION CODER _ ADD LOADER OF CAMERA REEL OFFLOAD
; TIMER_DRIVE_GET_CAMERA
; -------------------------------------------------------------------
; FROM -- Thu 14-Jun-2018 17:15:27
; TO ---- Thu 14-Jun-2018 18:44:00 -- 1 HOUR 30 MINUTE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 007
; DFX REQUIRED WORKING PROPERLY AND EXTRA FOR IT
; -------------------------------------------------------------------
; FROM -- Wed 20-Jun-2018 14:50:48
; TO ---- Wed 20-Jun-2018 15:02:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 008 PROBLEM THIS PROCESS REQUIRED SUSPENDER HOGGER
; CPU IN SHORT QUICK BURSTS
; -------------------------------------------------------------------
; Intel(R) Dynamic Platform and Thermal Framework Utility Application
; PROBLEM KEEP REPEAT RUNNING FOR SHORT SPACE FREQUENTLY AND REDUCE CPU COST
; Intel Corporation
; C:\Windows\Temp\DPTF\esif_assist_64.exe
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; STOP THE PROCESS WITH SUSPEND FOR AN HOUR AND LET BREATHE AGAIN 
; FOR A FEW MINUTE AND LOOPER
; [ Sunday 19:28:30 Pm_07 October 2018 ]
; -------------------------------------------------------------------
; DFX REQUIRED WORKING PROPERLY AND EXTRA FOR IT
; -------------------------------------------------------------------
; FROM -- Sun 07-Oct-2018 18:00:41
; TO ---- Sun 07-Oct-2018 19:24:00 _ 1 & HALF HOUR ALMOST _ INTO AN ALLNIGHTER AGAIN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 008 
; ADD CODE THIS ROUTINE _ GITHUB_MIDNIGHT_AND_MIDDAY_TIMER
; A NEW LEVEL OF TIMER PERIODIC ON THE HOUR OR MINUTE OR DAY
; ACTUALLY ON THE HOUR OR DAY MIDNIGHT SORT OF THING
; EASY DONE PROGRAMMED
; THAT ALL MY WORK WANTED TODAY
; COMPLETELY KNACKER-ED _ EASY AN ALLNIGHTER
; STILL NEXT PROJECT BEEN WAITING FOR 
; MY APK'S PACKS ON MOBILE WANT BACKING UP NICELY
; FILENAME RECURSIVE AND MOVE _.VBS
; GETTING A DAB HAND AT THEM REMEMBER WHEN SO HARD TO GET FILES 
; AND FOLDER TOGETHER
; AN THEN LOADS OF PHOTO'S TO SORT _ GETTING ON WITH IT QUICKER NOW
; I BEGAN PROGRAM IT PROPER
; -------------------------------------------------------------------
; FROM -- Sat 20-Oct-2018 06:29:02
; TO ---- Sat 20-Oct-2018 08:00:02 _ 1 AND HALF HOUR
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 009
; ADD CODE PIXEL FIND ROUTINE TO CHROME AND YAHOO PRINTER OUTPUT NEW THING
; ON PRESS P IN YAHOO - BIT BUGGY THEY HAVING PRINT ONLY COME UP ONCE SOMETIMES
; -------------------------------------------------------------------
; FROM -- Mon 22-Oct-2018 19:18:41
; TO ---- Mon 22-Oct-2018 20:12:14 - 1 HOUR 
; -------------------------------------------------------------------



;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 19-SCRIPT_TIMER_UTIL.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL.ahk
----
;# ------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
;--------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;--------------------

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

; DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

Set_Var_Responding_1=FALSE
Set_Var_Responding_2=FALSE
O_HWND_1=0
O_Status_GSDATA=1
SCRIPT_NAME_VAR=

GLOBAL VAR_STORE_CAMERA_LABEL
GLOBAL Secs_CAMERA
GLOBAL Label_CAMERA
GLOBAL OLDPID_ESIF_ASSIST_64_SUSPEND
GLOBAL SET_GO_ESIF_ASSIST_64_SUSPEND
GLOBAL HWND_ID_1_OLD

GLOBAL ID_OLD_ConsoleWindowClass
GLOBAL ID_ConsoleWindowClass_TIMER

GLOBAL ID_TeamViewer_Panel_TV_ControlWin_TIMER
GLOBAL ID_OLD_TeamViewer_Panel_TV_ControlWin

GLOBAL dhw

OLD_UniqueID_CHROME=0
OLD_UniqueID_RfEditor=0
OLD_UniqueID_NOTEPAD_PLUS_PLUS=0

ID_TeamViewer_Panel_TV_ControlWin_TIMER=0
ID_OLD_TeamViewer_Panel_TV_ControlWin=0

ID_OLD_ConsoleWindowClass=0
ID_ConsoleWindowClass_TIMER=0

HWND_ID_1_OLD=0

OLDPID_ESIF_ASSIST_64_SUSPEND=0
SET_GO_ESIF_ASSIST_64_SUSPEND=0

VAR_STORE_CAMERA_LABEL=

TIMER_SUB_OWNER_SAVE_TIMER=0

; Each array must be initialized before use:
COMPUTER_NAME_ARRAY := []

UniqueID_Old=0

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6
	

SCRIPT_NAME_VAR:=SubStr(A_ScriptName, 1, -4)
SCRIPT_NAME_VAR=%A_ScriptDir%\%SCRIPT_NAME_VAR%_TIMER_%A_ComputerName%.txt
SCRIPT_NAME_VAR=%SCRIPT_NAME_VAR%

If FileExist(SCRIPT_NAME_VAR)
{
	FileReadLine, TIMER_SUB_OWNER_SAVE_TIMER, %SCRIPT_NAME_VAR%, 1
}
	

setTimer TIMER_SUB_1,200

setTimer TIMER_SUB_EliteSpy, OFF

setTimer TIMER_SUB_GOODSYNC,1000
setTimer TIMER_SUB_GOODSYNC_OPTIONS,1000

setTimer TIMER_SUB_NOTEPAD_PLUS_PLUS,1000

setTimer TIMER_SUB_WINDOWS_UPDATE,OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_WINDOWS_DESKTOP_ICON,10000
; STARTS AS 10 SECOND AND THEN GOES TO 10 MINUTE

setTimer TIMER_SUB_VICE_VERSA,OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_SCRIPT_SHELL_FOLDERING,1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

;setTimer TIMER_SUB_OWNER, % -1 * 1000 * 60 * 60 ; After1Hours
setTimer TIMER_SUB_OWNER, 1000 ; After1Hours
; STARTS AS AFTER 1 HOUR AND THEN GOES TO EVERY 24 HOUR

;setTimer TIMER_SUB_WSCRIPT,100

;setTimer TIMER_SUB_CMD_KILL, OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_LOGGER, 1000
; LOGGING BLUETOOTH AND DUPLICATE CLEANER LOGGER TRUNCATE
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB__SendSMTP__0__LOG_BAT,1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE,10000
; STARTS AS 10 SECOND AND THEN GOES TO EVERY HALF HOUR

setTimer TIMER_SUB__MY_IP, 10000
; STARTS AS 10 SECOND AND THEN GOES TO EVERY 10 MINUTE

setTimer TIMER_PREVIOUS_INSTANCE, 10
; STARTS AS 10 MILLI-SECOND AND THEN GOES TO EVERY 1 MINUTE
; -------------------------------------------------------------------
; I used PREVIOUS_INSTANCE Detection For My Scripts Now 
; Otherwise they Maybe Two or More Loaded When You Run a Whole Bunch Together Quickly at Bootup or Just to run Again after Killed the Lot or Something
; Sometimes On Less Quicker Machine Computer
; There Might Be a Bug with This at Display the Icon in System Tray Not Sure Yet Maybe it was at a Time when My System was Instability
; -------------------------------------------------------------------


SETTIMER TIMER_DRIVE_CAMERA_UPLOAD_DROPBOX,4000

SETTIMER TIMER_SUB_ESIF_ASSIST_64_SUSPEND, 20000 ; ---- 20 SECONDS
SETTIMER TIMER_SUB_ESIF_ASSIST_64_SUSPEND_WAIT_AN_HOUR,3600000 ; ---- 1 HOUR


SETTIMER TIMER_Check_Any_PID_Suspended_Warning, 10000 ; ---- 10 SECONDS ---- And Then 1 Hour

GITHUB_MIDNIGHT_AND_MIDDAY_TIMER_DONE=
GOSUB GITHUB_MIDNIGHT_AND_MIDDAY_TIMER

RETURN


GITHUB_MIDNIGHT_AND_MIDDAY_TIMER:

	; ---------------------------------------------------------------
	; 1 = DAY TIMER 
	; 2 = HOUR TIMER
	; 3 = MINUTE TIMER
	; ---------------------------------------------------------------
	VALUE_TIMER_DY_HR_MI=2
	
	IF VALUE_TIMER_DY_HR_MI=1
	{
		Midnight := SubStr( A_Now, 1, 8 ) . "000000"
		Midnight += 1, days
	}
	IF VALUE_TIMER_DY_HR_MI=2
	{
		Midnight := SubStr( A_Now, 1, 10 ) . "000000"
		Midnight += 1, hours
	}
	IF VALUE_TIMER_DY_HR_MI=3
	{
		Midnight := SubStr( A_Now, 1, 12 ) . "000000"
		Midnight += 1, MINUTES
	}

	Midnight -= A_Now, seconds

	EnvMult, Midnight, 1000

	SET_GO=FALSE
	IF A_Hour=12
		SET_GO=TRUE
	IF A_Hour=0
		SET_GO=TRUE
	
	IF SET_GO=FALSE 
		RETURN
	
	IF GITHUB_MIDNIGHT_AND_MIDDAY_TIMER_DONE
	{
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.exe"
		IfExist, %FN_VAR%
		{
			Run, %FN_VAR% /GOODSYNC_MODE
		}
	}
	
	GITHUB_MIDNIGHT_AND_MIDDAY_TIMER_DONE=TRUE
	
	SETTIMER GITHUB_MIDNIGHT_AND_MIDDAY_TIMER, OFF
	SETTIMER GITHUB_MIDNIGHT_AND_MIDDAY_TIMER, %Midnight%
	SETTIMER GITHUB_MIDNIGHT_AND_MIDDAY_TIMER, ON
	; ----
	; Test Timer Status - Ask for Help - AutoHotkey Community
	; https://autohotkey.com/board/topic/55321-test-timer-status/
	; ----
RETURN


;----------------------------------------
;----------------------------------------
; BEGIN THE MAIN CODE
; CHOW MEIN
; MAIN CODE
;----------------------------------------
;----------------------------------------
;----------------------------------------
;----------------------------------------
TIMER_SUB_1:
	
DetectHiddenWindows, ON
SetTitleMatchMode 2

WinGet, ID, list,ahk_class ConsoleWindowClass

; Administrator:  C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 BOOT KILLER.BAT

IF ID_OLD_ConsoleWindowClass<>%ID%
{
	ID_ConsoleWindowClass_TIMER=%A_Now%
	;ID_ConsoleWindowClass_TIMER+=10*60 ; 10 MINUTES
	ID_ConsoleWindowClass_TIMER+=40 ; 40 SECOND
}

ID_OLD_ConsoleWindowClass=%ID%

; ------
; IF (ID)
; ------
If (A_Now<ID_ConsoleWindowClass_TIMER)
{
	Process, Exist, TASKLIST.EXE
	NewPID = %ErrorLevel%  ; Save the value immediately ErrorLevel is often changed
	If NewPID > 0
	{
		SoundBeep , 3000 , 100
		Process, priority, %NewPID%, Realtime
		ID_ConsoleWindowClass_TIMER=%A_Now%
		ID_ConsoleWindowClass_TIMER+=40 ; 40 SECOND
		}

	Process, Exist, TASKKILL.EXE
	NewPID = %ErrorLevel%  ; Save the value immediately ErrorLevel is often changed
	If NewPID > 0
	{
		SoundBeep , 2000 , 100
		Process, priority, %NewPID%, Realtime
		ID_ConsoleWindowClass_TIMER=%A_Now%
		ID_ConsoleWindowClass_TIMER+=40 ; 40 SECOND
	}
}


; MAYBE WANT IT
; DetectHiddenWindows, OFF


if (WinExist("Output ahk_class #32770"))
{
WinGet, HWND, ID, Output ahk_class #32770
WinGet, PATH, ProcessName, ahk_id %HWND%
IF PATH=ViceVersa.exe
	WINCLOSE
}


; -------------------------------------------------------------------
; [ Sunday 22:34:10 Pm_15 April 2018 ]
; Time Spend 
; Sun 15-Apr-2018 20:38:30 <> 
; Sun 15-Apr-2018 22:38:00 = 2 Hours
; I Frequently Want to swap between show Hidden Files in the System
; But it comes up with a Nagg Screen Warning
; But It Doesn't Anymore
; Question With The Answer _ Answer Sewer
; ---------------------------------------------------
; Hide Protected Operating System Files (Recommended)
; ---------------------------------------------------
; I wasn't Able to Make it Check Against Any Text Other than Warning
; Not any Control Text or Other Windows Text
; But control Copy Gets the Text for The Clipboard 
; There Was the Ability to Check if a Yes No Question was on the MsgBox
; But I Had to Do the Best I Able Do to Safeguard Hitting Another Window
; Project with More Subroutine Set
; Located Here Folder and Source File
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
; https://www.dropbox.com/s/yuspt7npqj0jpuo/Autokey%20--%2019-SCRIPT_TIMER_UTIL.ahk?dl=0
; -------------------------------------------------------------------

; MAYBE WANT IT
; DetectHiddenWindows, OFF

if (WinExist("Warning ahk_class #32770"))
{
	; Test Line Remmed
	; WinGet, OutputVar, ControlList, ahk_id %HWND%
	; Tooltip, % OutputVar

	WinGet, HWND, ID, Warning ahk_class #32770
	WinGetClass, This_Class, ahk_id %HWND%
	WinGet, path, ProcessName, ahk_id %HWND%
	WinGetText, OutputVar, ahk_id %HWND%
	WinGetPos, WinLeft, WinTop, WinWidth, WinHeight, ahk_id %HWND%

	IF (OSVER_N_VAR=5) 
	{
		WinWidthOS=WinWidth
		WinHeightOS=WinHeight
	}
	IF (OSVER_N_VAR=6)  
	{
		WinWidthOS=572
		WinHeightOS=201
	}
	IF (OSVER_N_VAR=10)
	{
		WinWidthOS=713
		WinHeightOS=256
	}
	
	; Include Multiple If-And Statement With Separated Lines
	; ---------------------------------
	IF (This_Class="#32770" 
	and PATH="explorer.exe" 
	and OutputVar="&Yes`r`n&No`r`n" 
	and WinWidth=WinWidthOS
	and WinHeight=WinHeightOS)
	{
		Control, Check,, Button1
		SoundBeep , 4000 , 100
	}
}

DetectHiddenWindows, OFF

;Would you like to switch to the following audio playback device?
IfWinExist FxSound Message
{
	;WinGet, OutputVar, ControlList, FxSound Message
	;Tooltip, % OutputVar ; List All Controls of Active Window
	;---------------------------------------------------------
	ControlGettext, OutputVar_2, Static1,  FxSound Message
	
	IfInString, OutputVar_2, Would you like to
	{
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
	    Control, Check,, Button4, FxSound Message ahk_class #32770
	    ControlClick, Button2, FxSound Message ahk_class #32770
	}
}

IfWinExist File Access Denied ahk_class #32770
{
	;WinGet, OutputVar, ControlList, FxSound Message
	;Tooltip, % OutputVar ; List All Controls of Active Window
	;---------------------------------------------------------
	ControlGettext, OutputVar_2, SysLink1, File Access Denied ahk_class #32770
	
	; MSGBOX % OutputVar_2
	
	IfInString, OutputVar_2, You'll need to provide administrator
	{
	    ControlClick, Button1, File Access Denied ahk_class #32770
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
	}
}

IfWinExist File Access Denied ahk_class OperationStatusWindow
{

	WINACTIVATE, File Access Denied ahk_class OperationStatusWindow
	SENDINPUT {ENTER}
	
	; ControlClick, Continue, File Access Denied ahk_class OperationStatusWindow
	SoundBeep , 3000 , 100
	SoundBeep , 2000 , 100
}

DetectHiddenWindows, ON

ifWinNotExist, ahk_class wndclass_desked_gsk
{
	if (WinExist("(Not Responding)") and Set_Var_Responding_1="FALSE")
	{
		Set_Var_Responding_1=TRUE
		SoundBeep , 1000 , 100
		SoundBeep , 2000 , 100
	}

	if (!WinExist("(Not Responding)") and Set_Var_Responding_1="TRUE")
	{
		Set_Var_Responding_1=FALSE
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
		SoundBeep , 2500 , 100
		SoundBeep , 2000 , 100
	}
}

if (WinExist("Page Unresponsive") and Set_Var_Responding_2="FALSE")
{
	Set_Var_Responding_2=TRUE
	SoundBeep , 1000 , 100
	SoundBeep , 2000 , 100
}

if (!WinExist("Page Unresponsive") and Set_Var_Responding_2="TRUE")
{
	Set_Var_Responding_2=FALSE
	SoundBeep , 3000 , 100
	SoundBeep , 2000 , 100
	SoundBeep , 2500 , 100
	SoundBeep , 2000 , 100
}

DetectHiddenWindows, OFF

IfWinExist RoboForm Upgrade
{
	WINCLOSE
	;WinActivate
	;sendinput, !{F4}		I
	;SoundBeep , 2500 , 100
}

IfWinExist ahk_class #32770
{

	HWND_ID_1 := WinExist("ahk_class #32770")
	HWND_ID_2 := WinExist("AutoSave Login - RoboForm ahk_class #32770")
	HWND_ID_3 := WinExist("Save - RoboForm ahk_class #32770")
	HWND_ID_4 := WinExist("Install RoboForm ahk_class #32770")
	HWND_ID_5 := WinExist("Sync RoboForm Data Folder")
	HWND_ID_6 := WinExist("AutoFill - RoboForm")
	HWND_ID_7 := WinExist("AutoSave New Account - RoboForm")
	
	ControlGetText, OutputVar_1, _RoboForm_Dialog_1100973_, ahk_class #32770
	
	; -------------------------------------------------------
	; THIS IS THE THIN TASK BAR UNDERNEATH THE APP'S
	; NOT MUCH TO GO ON TAKEN MAX THOUGH _ FILTER THIS ONE TO 
	; IGNORE TO ACT ON
	; MORE EFFORT ONLY ABLE USE DON'T SAVE WORD
	; -------------------------------------------------------
	; Save
	; Don't Save
	; _RoboForm_Dialog_1100973_
	; -------------------------------------------------------
	ControlGetText, OutputVar_2, Don't Save, ahk_class #32770
	; ControlGetText, OutputVar_2, Save`r`nDon't Save`r`n_RoboForm_Dialog_1100973_, ahk_class #32770
	
	; Fill Empty Fields &Only
	; _RoboForm_Dialog_1100973_
	; -------------------------------------------------------
	ControlGetText, OutputVar_3, Fill Empty Fields, ahk_class #32770
	
	SET_GO=FALSE
	IF OutputVar_1
		SET_GO=TRUE
	IF OutputVar_2
		SET_GO=FALSE
	IF OutputVar_3
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_2%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_3%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_4%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_5%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_6%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_7%
		SET_GO=FALSE
		
	IF SET_GO=TRUE
	{	
		SoundBeep , 2500 , 100
		;MSGBOX REQUIRE HELP HERE ROBOFORM SEARCHING _RoboForm_Dialog_1100973_
		;ControlClick, No, ahk_class #32770
	}
}


DetectHiddenWindows, On

; -------------------------------------------------------------------
; Question at End of Sync is Able to Be Made Close Auto
; -------------------------------------------------------------------
Sting_Var=Sync RoboForm Data Folder

IfWinExist %Sting_Var%
{
	HWND_ID_5 := WinExist("%Sting_Var%")
	SET_GO=FALSE
	; ---------------------------------------------------------------
	; ON SYSTEM START UP OF 1ST RUN _ Button60
	; ---------------------------------------------------------------
	ControlGetText, OutputVar, Button60, %Sting_Var%
	IfInString, OutputVar, &Close
		SET_GO=TRUE
	; ---------------------------------------------------------------
	; ON REQUEST MANUAL SYNC        _ Button40
	; ---------------------------------------------------------------
	ControlGetText, OutputVar, Button40, %Sting_Var%
	IfInString, OutputVar, &Close
		SET_GO=TRUE

	IF SET_GO=TRUE 
	{
		DetectHiddenText, Off
		WinGetText, OutputVar_1, %Sting_Var%
		if OutputVar_1=&Close`r`n
			ControlClick, &Close, %Sting_Var%
	}
}

; -----------------------
; Default Setting By Here
; -----------------------
DetectHiddenText, On

IfWinExist RoboForm Update ahk_class #32770
{
	SoundBeep , 2500 , 100
	ControlClick, No, RoboForm Update ahk_class #32770
}

IfWinExist ahk_class #32770
{
	ControlGetText, OutputVar, Button1, ahk_class #32770
	IfInString, OutputVar, Install
	{	
	WINCLOSE ahk_class #32770 ahk_exe robotaskbaricon.exe
	;ControlClick, Button1, ahk_class #32770
	}
}

;This operation will affect the entire list, not just the current filter. Continue?
IfWinExist DuplicateCleaner ahk_class #32770
{
	WinGetText, OutputVar, DuplicateCleaner ahk_class #32770
	IfInString, OutputVar, This operation will affect the entire list
	{	
	ControlClick, Button1, DuplicateCleaner ahk_class #32770
	}
}


; DetectHiddenWindows, on

;6 days 22 hours left in trial
;ahk_class #32770
;ahk_exe dfx.exe

DetectHiddenWindows, off

IfWinExist left in trial
{
	WinActivate
	sendinput, !{F4}		I
	SoundBeep , 1000 , 50
}

SetTitleMatchMode 3  ; Exactly

; DFXSTATIC
; ..\..\UTIL\doubleBuff\doubleBuffInit.cpp
IfWinExist 38 ahk_class #32770
{
	; ControlGetText, OutputVar, doubleBuffInit , 38 ahk_class #32770
	; ControlGetText,Static1,doubleBuffInit
	; ControlGettext, OutputVar, Static1, 38 ahk_class #32770
	; MSGBOX % OutputVar 
	; IF OutputVar 

	ControlClick, OK, 38 ahk_class #32770
	SoundBeep , 1000 , 50
}
IfWinExist 58 ahk_class #32770
{
	ControlClick, OK, 58 ahk_class #32770
	SoundBeep , 1000 , 50
}
IfWinExist 158 ahk_class #32770
{
	ControlClick, OK, 158 ahk_class #32770
	SoundBeep , 1000 , 50
}

SetTitleMatchMode 2

;---------------------------------------
;---------------------------------------
;____ TEAM_VIEWER

DetectHiddenWindows, on

IfWinExist Sponsored session
{
	;WinActivate
	;sendinput, !{F4}		I
	SoundBeep , 2500 , 100
	
	;0x10 CLOSE
	;PostMessage, 0x10, 0, 0,, Sponsored session
	
	;PostMessage, 0x112, 0xF060,,, WinTitle, WinText  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
	;PostMessage, 0x112, 0xF060,,, ,Sponsored session,   ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE

	ControlClick, OK, Sponsored session
}
;--------------------------------------------------------------------
; TEAM VIEWER CLOSE AFTER CONNECTION ADVERT
;--------------------------------------------------------------------

DetectHiddenWindows, OFF

IfWinExist ahk_class ATL:03A50E50
{
	ControlClick, OK, ahk_class ATL:03A50E50
}

IfWinExist Commercial use ahk_class #32770
{
	;ControlClick, Button4, Commercial use ahk_class #32770
	ControlClick, OK, Commercial use ahk_class #32770
}


IfWinExist Session timeout! ahk_class #32770
{
	ControlClick, OK, Session timeout! ahk_class #32770
}



DetectHiddenWindows, OFF
SetTitleMatchMode 3  ; Exactly
;Visual BASIC
IfWinExist Visual Component Manager
{
	ControlClick, OK, Visual Component Manager
	SoundBeep , 2500 , 100
}

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.
DetectHiddenWindows, OFF
;Visual BASIC
IfWinExist Data View
{
	SoundBeep , 2500 , 100
	ControlClick, OK, Data View
}

DetectHiddenWindows, ON

IfWinExist hubiC
{
	ControlGetText, OutputVar, hubiC is already running. , hubiC
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, OK, hubiC
	}
}

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.
IfWinExist hubiC
{
	ControlGetText, OutputVar, If you exit hubiC now , hubiC
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, &Yes, hubiC
	}
}


; MAYBE WANT IT
DetectHiddenWindows, OFF

IfWinExist RoboForm Question
{
	ControlGetText, OutputVar, Do you want to delete
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, OK, RoboForm Question
	}
}

; MAYBE WANT IT
; DetectHiddenWindows, OFF

IfWinExist Microsoft OneDrive
{
	ControlGetText, OutputVar, Close OneDrive , Microsoft OneDrive
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, Close OneDrive, Microsoft OneDrive
	}
}

; MAYBE WANT IT
; DetectHiddenWindows, OFF

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinNotExist, Replace
	UniqueID_Old=0

IfWinExist, Replace
{
	; Replace Search and Replace Move to Better Position In Visual Basic
	; or any other editor
	ControlGetText, OutputVar, Current &Procedure , Replace
	IF OutputVar 
	{	
		UniqueID := WinExist("Replace")
		;tooltip %UniqueID%
		WinGetPos,,YPos,,, Replace
		if (YPOS>(A_ScreenHeight/2))
			UniqueID_Old=0
		if UniqueID_Old<>%UniqueID%
		{
			;WinMove, Replace,, (A_ScreenWidth/2)-(Width/2), (A_ScreenHeight/2)-(Height/2)
			SoundBeep , 2500 , 50
			WinGetPos,,YPos, Width, Height, Replace
			WinMove, Replace,, (A_ScreenWidth)-(Width), 0

			WinGetPos,,YPos, Width, Height, Replace
			if (YPOS<>0)
				UniqueID=0
		}
		
		UniqueID_Old=%UniqueID%
	}
}


; MAYBE WANT IT
; DetectHiddenWindows, OFF

;DuplicateCleaner
IfWinExist Finished deleting files.
{
	SoundBeep , 2500 , 100
	ControlClick, OK, Finished deleting files.
}

; MAYBE WANT IT
; DetectHiddenWindows, OFF

;Are you sure you want to cancel the scan?
IfWinExist DuplicateCleaner
{
	ControlGetText, OutputVar, Are you sure you want to cancel , DuplicateCleaner
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, &Yes, DuplicateCleaner
	}
}


; --------------------------------------------------------------------
; Stopping With Warning About Open a Batchfile in Scripter GitHub Folder
; --------------------------------------------------------------------
; C:\Windows\SystemApps\Microsoft.Windows.AppRep.ChxApp_cw5n1h2txyewy\CHXSmartScreen.exe
; HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System
; DWord
; EnableSmartScreen=0
; --------------------------------------------------------------------
; This one worked Instant Change to change the Form From a Windows 10 APP to a Normal Form Window
; HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer
; String
; SmartScreenEnabled
; Off
; --------------------------------------------------------------------
; The publisher could not be verified. Are you sure that you want to run this software?
; Open File - Security Warning
; --------------------------------------------------------------------
; ----
; Change Windows SmartScreen Settings in Windows 10 | Windows 10 Tutorials
; https://www.tenforums.com/tutorials/5593-change-windows-smartscreen-settings-windows-10-a.html
; --------
; Turn On or Off SmartScreen for Windows Store Apps in Windows 10 | Windows 10 Tutorials
; https://www.tenforums.com/tutorials/81139-turn-off-smartscreen-windows-store-apps-windows-10-a.html
; ----

; --------------------------------------------------------------------
; I TRY AND SET THE REGKEY AUTO ON START UP \Autokey -- 21-AutoRun.ahk
; BUT NOT ALLOW PERMISSION ONLY ALLOWED IN REGEDIT MANUALLY
; --------------------------------------------------------------------
; BUT I CAN'T WRITE THIS REGISTRY KEY WITHOUT GOING INTO REGEDIT 
; SOMEHOW NOT PERMISSION
; --------------------------------------------------------------------
; RegWrite, REG_SZ, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer, SmartScreenEnabled, Off
; --------------------------------------------------------------------

; MAYBE WANT IT
; DetectHiddenWindows, OFF

IfWinExist Open File - Security Warning
{
	ControlGetText, OutputVar, The publisher could not be verified , Open File - Security Warning
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, &Run, Open File - Security Warning
	}
}

; MAYBE WANT IT
; DetectHiddenWindows, OFF

;DuplicateCleaner
IfWinExist Scan cancelled
{
	SoundBeep , 2500 , 100
	ControlClick, Close, Scan cancelled
}



; MAYBE WANT IT
; DetectHiddenWindows, OFF

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinExist MSDN Library Visual
{
	ControlGetText, OutputVar, Active Su&bset , MSDN Library Visual
	IF OutputVar 
	{	
		;Esc::
		;{
		;	SoundBeep , 3500 , 100
		;	PostMessage, 0x112, 0xF060,,, MSDN Library Visual  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
		;}
		;return
	}
}


;----
;Joystick Test Script
;https://autohotkey.com/docs/scripts/JoystickTest.htm
;----

;----
;Detect if a variable is empty/blank - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?t=5080
;----
	

; -------------------------------------------------------------------
; C:\Program Files (x86)\TeamViewer\TeamViewer.exe
; AUTOFILL PASSWORD DIFFICULT
; AND FREQUENT REMINDER
; -------------------------------------------------------------------

; MAYBE WANT IT
; DetectHiddenWindows, OFF

SetTitleMatchMode 3  ; Exactly

IfWinExist Configure Permanent Access
{
	HWND_ID_1 := WinExist("Configure Permanent Access")

	IF HWND_ID_1<>%HWND_ID_1_OLD%
	{
		HWND_ID_1_OLD=%HWND_ID_1%
		
		SoundBeep , 2500 , 100
		; ControlClick, &Run, Open File - Security Warning
		
		FN_VAR:="C:\RF\7-ASUS-GL522VW\LOGIN APPS 01\TeamViewer - EXE APP.rfp"
		IfExist, %FN_VAR%
		{
			Run, %FN_VAR%
		}
	}
}
IfWinExist Permanent Access Activated
{
	SoundBeep , 2500 , 100
	ControlClick, OK, Permanent Access Activated
}

SetTitleMatchMode 3  ; Exactly

; CHROME
DetectHiddenText, ON
UniqueID := WinActive("ahk_class Chrome_WidgetWin_1")
IF UniqueID>0 
IF OLD_UniqueID_CHROME<>%UniqueID%
{
	SET_GO=TRUE
	; ---------------------------------------------------------------
	; SOME EXTENSION THAT LOOK LIKE TOOL-TIPS DON;T HAVE THE ' b ' IN THE INFO
	; FILTER THEM OUT OR BE MAXIMIZE IT ALL
	; ---------------------------------------------------------------
	; b
	; Chrome Legacy Window
	; ---------------------------------------------------------------
	WinGetText, OutputVar, ahk_id %UniqueID%
	IfNotInString, OutputVar, Chrome Legacy Window`r`n
		SET_GO=FALSE

	LOOP, 4
	{
		IF SET_GO=TRUE
			SLEEP 100
		WinGetText, OutputVar, ahk_id %UniqueID%
		IfNotInString, OutputVar, b`r`n
				SET_GO=FALSE
		IF SET_GO=FALSE
				BREAK
	}
		
		
	IF SET_GO=TRUE
	{	
		WinMaximize, ahk_id %UniqueID%
		SoundBeep , 2500 , 100
	}
	OLD_UniqueID_CHROME=%UniqueID%
}

DetectHiddenText, OFF

SetTitleMatchMode 3  ; Exactly
	
; ROBOFORM EDITOR
UniqueID := WinActive("ahk_class RfEditor")
IF UniqueID>0 
IF OLD_UniqueID_RfEditor<>%UniqueID%
{
	OLD_UniqueID_RfEditor=%UniqueID%
    WinMaximize  ; Maximizes the Notepad window found by IfWinActive above.
	SoundBeep , 2500 , 100
}
	
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.
; Notepad++
UniqueID := WinActive(" - Notepad++")
IF UniqueID>0 
IF OLD_UniqueID_NOTEPAD_PLUS_PLUS<>%UniqueID%
{
	OLD_UniqueID_NOTEPAD_PLUS_PLUS=%UniqueID%
    WinMaximize  ; Maximizes the Notepad window found by IfWinActive above.
	SoundBeep , 2500 , 100
}
	

; -------------------------------------------------------------------
; IF NOT WANT TO CLICK ON __ AutoFill - RoboForm __ TO GAIN 
; FOCUS AND OPERATE
; INSTEAD GOT TO GO TO ROBOFORM SETTING AND CLICK 
; AUTOFILL STEALS KEYBOARD FOCUS
; -------------------------------------------------------------------
; -------------------------------------------------------------------
SetTitleMatchMode 3  ; Exactly

WINDOW_Array := []

ArrayCount := 0
ArrayCount += 1
WINDOW_Array[ArrayCount] := "Email Login Page - Google Chrome"	
ArrayCount += 1
WINDOW_Array[ArrayCount] := "NAS-QNAP-ML - Google Chrome"

Loop % ArrayCount
{
  Element := WINDOW_Array[A_Index]
	UniqueID := WinActive("AutoFill - RoboForm")
	IF UniqueID>0 
		IfWinExist %Element%
		{
			ControlGettext, OutputVar_2, Button1, AutoFill - RoboForm
			If (OutputVar_2="&Fill Forms")
			{
				#WinActivateForce, AutoFill - RoboForm
				ControlClick, Button1, AutoFill - RoboForm
				SoundBeep , 2500 , 100
				#WinActivateForce, %Element%
				
				Loop, 30
				{
					IfWinExist %Element%
					{
						SLEEP 500
						SENDINPUT {ENTER}
						SoundBeep , 2500 , 100
					}
					ELSE
						BREAK
				}
				
			}
		}
}
	
	

SetTitleMatchMode 3  ; Exactly

UniqueID := WinActive("Log in with PayPal - Google Chrome")
IF UniqueID>0 
	
	Loop, 30
	{
		IfWinExist ahk_id %UniqueID%
		{
			#WinActivateForce, ahk_id %UniqueID%
			SLEEP 500
			SENDINPUT {ENTER}
			SoundBeep , 2500 , 100
		}
	}
	

SetTitleMatchMode 3  ; Exactly
DetectHiddenText, Off
IfWinExist TeamViewer ahk_class #32770
{
	WinGetText, OutputVar, TeamViewer ahk_class #32770
	IfInString, OutputVar, Show running TeamViewer
	{	
		ControlClick, Button3, TeamViewer ahk_class #32770
		SoundBeep , 2500 , 100

	}
}

; TEAMVIEWER ADVERT AT END OF SESSION
; -----------------------------------
; ahk_class ATL:03259F10
; ahk_exe TeamViewer.exe
IfWinExist ahk_class ATL:03259F10
{
	#WinActivateForce, ahk_class ATL:03259F10
	WINCLOSE ahk_class ATL:03259F10
}

SetTitleMatchMode 3  ; Exactly
DetectHiddenText, ON
; UniqueID := WinActive("TeamViewer ahk_class #32770")
; UniqueID := WinGET("TeamViewer ahk_class #32770")
WinGet, UniqueID, ID, TeamViewer ahk_class #32770

IF UniqueID>0 
{
	; MSGBOX % UniqueID
	WinGetText, OutputVar, ahk_id %UniqueID%
	; MSGBOX %  OutputVar
	SET_GO=FALSE
	IfInString, OutputVar, Loading...
		SET_GO=TRUE
	IfInString, OutputVar, Please wait
		SET_GO=TRUE
	IF SET_GO=TRUE
	{	
		#WinActivateForce, ahk_id %UniqueID%
		Loop, 30
		{
			IfWinExist ahk_id %UniqueID%
			{
				SLEEP 500
				WINCLOSE ahk_id %UniqueID%
				SoundBeep , 2500 , 100
			}
		}
		
	}
; Loading...
; Please wait
; 2
; 3
; 4
; 5
}


SetTitleMatchMode 3  ; Exactly
DetectHiddenText, Off
HWND_ID_1 := WinExist(".NET-BroadcastEventWindow.4.0.0.0.1a8c1fa.0: chrome.exe - Application Error")
IF HWND_ID_1>0
{
	ControlClick, OK, ahk_id %HWND_ID_1%
	SoundBeep , 2500 , 100
}


SetTitleMatchMode 3  ; Exactly
DetectHiddenText, Off
HWND_ID_1 := WinExist("FUJIFILM PC AutoSave")
IF HWND_ID_1>0
{
	ControlClick, OK, ahk_id %HWND_ID_1%
	SoundBeep , 2500 , 100
}


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

SetTitleMatchMode 3  ; Exactly
DetectHiddenText, Off
HWND_ID_1 := WinExist("InfoRapid Search & Replace ahk_class #32770")
IF HWND_ID_1>0
{
	ControlClick, &No, ahk_id %HWND_ID_1%
	SoundBeep , 2500 , 100
}

	
Return
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; END OF TIMER_1
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
TIMER_DRIVE_CAMERA_UPLOAD_DROPBOX:
; -------------------------------------------------------------------
IfWinExist Camera Upload
{
	WinActivate
	sendinput, !{F4}		; CLOSE
	SoundBeep , 1000 , 50
}
RETURN

; -------------------------------------------------------------------
TIMER_SUB_WINDOWS_DESKTOP_ICON:
; -------------------------------------------------------------------
; THIS ROUTINE WONT DELETE ANYTHING IF CALLED WITHIN A SECOND OF THE PROGRAM STARTING 
; ONLY AFTER TIME WHEN LOPP COME AROUND AGAIN
; -------------------------------------------------------------------

setTimer TIMER_SUB_WINDOWS_DESKTOP_ICON,600000 ; 10 MINUTE

;C:\Windows10Upgrade\Windows10UpgraderApp.exe /ClientID "Win10Upgrade:VNL:NHV13SIH:{}"

SET_GO=FALSE
IF (A_ComputerName="4-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName="7-ASUS-GL522VW") 
	SET_GO=TRUE

IF SET_GO=TRUE
{
	FN_VAR="E:\01 Desktop\#_%A_ComputerName%\Windows 10 Update Assistant.lnk"
	FN_VAR:=StrReplace(FN_VAR, """" , "")
	FN_VAR_1=%FN_VAR%
	IfExist, %FN_VAR%
		{
			SoundBeep , 2000 , 200
			FileDelete, %FN_VAR%
		}
}	

; -------------------------------------------------------------------
; STRIP QUOTE WHY AS ABLE LEAVE THEM OUT IN FIRST PLACE
; IfExist DON'T WORK WITH QUOTES AROUND FILENAME
; -------------------------------------------------------------------
FN_VAR_1=E:\01 Desktop\#_%A_ComputerName%\GoodSync Explorer.lnk
FN_VAR_1:=StrReplace(FN_VAR_1, """" , "")
FN_VAR_2=%FN_VAR_1%
IfExist, %FN_VAR_1%
	{
		SoundBeep , 2000 , 200
		FileDelete, %FN_VAR_1%
	}

; LESS_TIMER=TRUE
; IfExist, %FN_VAR_1%
	; LESS_TIMER=FALSE
; IfExist, %FN_VAR_2%
	; LESS_TIMER=FALSE

; IF LESS_TIMER=TRUE
	; setTimer TIMER_SUB_WINDOWS_DESKTOP_ICON,600000 ; 10 MINUTE
	
RETURN


; -------------------------------------------------------------------
Multiple_Thread_Port_Scanner_ROUTINE:
; -------------------------------------------------------------------
; ---------------------------------------------------
; NOT USED IN CODE YET BUT PRACTICED FOR CODE BUILDER
; ---------------------------------------------------
; FROM    Thu 03-May-2018 16:16:39
; TO      Thu 03-May-2018 17:58:00 __ 1 HOUR 40 MINUTE
; ---------------------------------------------------
; I WROTE THIS CODE FROM IDEA WORKING ON CODE BEFORE
; AND SIDE TRACKED INTO THIS ONE
; AND NEAR THE END DECIDED NOT GOING TO USE IT
; BUT PRACTICE WAS HELPFUL
; ---------------------------------------------------

IfNotExist, %A_TEMP%\IPTEST.TXT
{
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\Multiple_Port_Scanner.exe"
IfExist, %FN_VAR%
	{
		Run %comspec% /c ""%FN_VAR%" "/ALL" "" >"%A_TEMP%\IPTEST.TXT , , MIN
	}
}

ArrayCount = 0
; Write to the array:
Loop, Read, %A_TEMP%\IPTEST.TXT
{
	SET_GO=FALSE
	IF SubStr(A_LoopReadLine, 1 , 2)="\\"
		SET_GO=TRUE
	
	IP_LINE_1=%A_LoopReadLine%
	IP_LINE_2:=SubStr(A_LoopReadLine, 3)
	IP_LINE_2:=StrReplace(IP_LINE_2, "-" , "_")
	;StringLower, IP_LINE_2, IP_LINE_2
	;StringLower, IP_LINE_1, IP_LINE_1
	
	IP_LINE_3:="_03_fat32_4gb"
	
	IF SET_GO=TRUE
		{
		SET_GO=FALSE
		ifExist, %IP_LINE_1%\%IP_LINE_2%%IP_LINE_3%\*.*
			{
			SET_GO=TRUE
			IP_LINE_4=%IP_LINE_1%\%IP_LINE_2%%IP_LINE_3%
			;MSGBOX % IP_LINE_4
			}
		}
			
	IF SET_GO=TRUE
	{
		ArrayCount += 1  
		COMPUTER_NAME_ARRAY%ArrayCount% = %IP_LINE_4%
	}
}

; Read from the array:
Loop %ArrayCount%
{
    MsgBox % "Element number " . A_Index . " is " . COMPUTER_NAME_ARRAY%A_Index%
}

RETURN


; -------------------------------------------------------------------
TIMER_SUB_NOTEPAD_PLUS_PLUS:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.
DetectHiddenWindows, % dhw

; Notepad++ v7.5.8 Setup
IfWinExist Notepad++ v
{
	ControlGetText, OutputVar, Allow plugins to be loaded from, Notepad++
	IF OutputVar 
	{	
		ControlGet, Status, Checked,, Button5
		If Status = 0
		{
			Control, Check,, Button5
			SoundBeep , 4000 , 100
		}
	}
}
Return

; -------------------------------------------------------------------
TIMER_SUB_EliteSpy:
; -------------------------------------------------------------------

dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist EliteSpy+ by Andrea
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
	IfExist, %FN_VAR%
		{
			Run, "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
		}
}
DetectHiddenWindows, % dhw
Return

; -------------------------------------------------------------------
TIMER_SUB_GOODSYNC_OPTIONS:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

WinGet, HWND_1, ID, ] Options ahk_class #32770
	IF (HWND_1>0)
	{
		; WinGet, OutputVar, ControlList, ahk_id %HWND_1%
		; Tooltip, % OutputVar ; List All Controls of Active Window
		;---------------------------------------------------------
		ControlGettext, OutputVar_2, Button16, ahk_id %HWND_1%

		ControlGet, OutputVar_1, Line, 1, Edit9, ahk_id %HWND_1%
		
		WinGetTitle OutputVar_3,ahk_id %HWND_1%
		
		If (OutputVar_1 = 2
			and OutputVar_2="Periodically (On Timer), every")
			{
				ControlSetText, Edit9,, ahk_id %HWND_1%
				Control, EditPaste, 4, Edit9, ahk_id %HWND_1%
				SoundBeep , 4000 , 100

		}
		if O_HWND_1<>%HWND_1%
		{
			ControlGet, OutputVar_4, Visible, , Button16, ahk_id %HWND_1%
			ControlGet, Status, Checked,, Button16, ahk_id %HWND_1%
			If Status=0
			{
				Control, Check,, Button16, ahk_id %HWND_1%
				IF OutputVar_4=1
					SoundBeep , 4000 , 100
			}
		}
			
		ControlGettext, OutputVar_2, Button20, ahk_id %HWND_1%
		ControlGet, OutputVar_1, Line, 1, Edit2, ahk_id %HWND_1%
		
		If (!OutputVar_1 
			and OutputVar_2="Wait for Locks to clear, minutes")
			{
				ControlSetText, Edit12,, ahk_id %HWND_1%
				Control, EditPaste, 10, Edit2, ahk_id %HWND_1%
				SoundBeep , 4000 , 100

		}
		ControlGet, Status, Checked,, Button20, ahk_id %HWND_1%
		If Status=1
		{
			Control, UnCheck,, Button22, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}

		;------------------------------------------------------------
		; Button5 ---- Save deleted/replaced files to History f (...)
		; Button6 ---- Cleanup _history_ folder after this many (...)
		; IF BUTTON 5 SET THEN ALSO SET BUTTON 6 _ DON'T LEAVE INFINITE
		;------------------------------------------------------------
		ControlGet, Status, Checked,, Button5, ahk_id %HWND_1%
		If Status = 1
		{
			ControlGet, Status, Checked,, Button6, ahk_id %HWND_1%
			If Status = 0
			{
			Control, Check,, Button6, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
			}
		}

		;------------------------------------------------------------
		; ---- HISTORY SET DAYS TO 30
		;------------------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 1
		{
			ControlGet, OutputVar_1, Line, 1, Edit2, ahk_id %HWND_1%
			ControlGet, OutputVar_4, Enabled, , Edit2, ahk_id %HWND_1%
			
			If (Trim(OutputVar_1)<>30 and OutputVar_4=1)
				{
					ControlSetText, Edit2,, ahk_id %HWND_1%
					Control, EditPaste, 30, Edit2, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
			}
		}
		}
		
		;-----------------------------------------------------
		; Button3
		; Text:	Save deleted/replaced files to Recycle B (...)
		; Button4
		; Text:	Cleanup _saved_ folder after this many d (...)
		;-----------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 1
		{
			sleep, 500
			ControlGet, OutputVar_1, Line, 1, Edit1, ahk_id %HWND_1%
			ControlGet, OutputVar_4, Visible, , Edit1, ahk_id %HWND_1%
			If (Trim(OutputVar_1)<>30 and OutputVar_4=1)
				{
					ControlSetText, Edit1,, ahk_id %HWND_1%
					Control, EditPaste, 30, Edit1, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
			}
		}
		}
		
		;-----------------------------------------------------
		; CAN'T DO THIS ONE 
		; IT'S EITHER HISTORY OR SAVED NOT BOTH
		; WELL YOU CAN DO BUTTON4 SAVE DO-ER INFINITE
		; Button3
		; Text:	Save deleted/replaced files to Recycle B (...)
		; Button4
		; Text:	Cleanup _saved_ folder after this many d (...)
		; Button6
		; Cleanup _history_ folder after this many (...)
		; BUTTON4 AND BUTTON6 CAN ALWAYS BE ON
		;-----------------------------------------------------
		;ControlGet, Status, Checked,, Button3, ahk_id %HWND_1%
		;If Status = 0
		;{
		;	Control, Check,, Button3, ahk_id %HWND_1%
		;	SoundBeep , 4000 , 100
		;}
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 0
		{
			Control, Check,, Button4, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
		ControlGet, Status, Checked,, Button6, ahk_id %HWND_1%
		If Status = 0
		{
			Control, Check,, Button6, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
		
		
		IF O_HWND_1<>%HWND_1%
		{
			O_Status_GSDATA=0
		}
		;------------------------------------------------------------
		; ---- No _gsdata_ folder here
		;------------------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
			ControlGet, Status, Checked,, Button41, ahk_id %HWND_1%
			If Status = 1
			IF O_Status_GSDATA<>%Status%
			{
				Control, unCheck,, Button41, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
				ControlGet, Status, Checked,, Button41, ahk_id %HWND_1%
				If Status=0
					O_Status_GSDATA=1
			}
		}
		
		;------------------------------------------------------------
		;Button10 ---	Text:	Exclude empty folders
		;Button11 --- Text:	Exclude Hidden files and folders
		;Button12 -- Text:	Exclude System files and folders
		;------------------------------------------------------------
		Var_check=[VB EXE SYNC
		if (SubStr(OutputVar_3, 1, 12)=Var_check)
		{
			ControlGet, Status, Checked,, Button10, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button10, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			ControlGet, Status, Checked,, Button11, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button11, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			ControlGet, Status, Checked,, Button12, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button12, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
		}
		
		
		
		O_HWND_1=%HWND_1%
		
}

;Are you sure you want to move _GSDATA_ folder back to the sync folder?
;This option is for Advanced users and you should read the manual first.
;Yes
;No

;WinGet, HWND_1, ID, GoodSync Warning ahk_class #32770
WinGet, HWND_1, ID, GoodSync ahk_class #32770
WinGetText OutputVar_3,ahk_id %HWND_1%
;tooltip % HWND_1

IfInString, OutputVar_3, Are you sure you want to move _GSDATA_
{
	#WinActivateForce, ahk_id %HWND_1%
	ControlClick, Button2,ahk_id %HWND_1%
	SoundBeep , 4000 , 100
}

IfInString, OutputVar_3, Are you sure you want to keep _GSDATA_
{
	#WinActivateForce, ahk_id %HWND_1%
	ControlClick, Button2,ahk_id %HWND_1%
	SoundBeep , 4000 , 100
}
DetectHiddenWindows, % dhw
Return

;--------------------------------------------------------------------
TIMER_SUB_GOODSYNC:
;--------------------------------------------------------------------
;setTimer TIMER_SUB_GOODSYNC, OFF
dhw := A_DetectHiddenWindows
DetectHiddenWindows, OFF
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IF (TRUE=FALSE)
{
	Process, Exist, GoodSync-v10.exe
	If Not ErrorLevel
		{
		FN_VAR:="C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
		IfExist, %FN_VAR%
			{
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			Run, "%FN_VAR%" , , MIN
			}
		}
}

IfWinExist GoodSync - Preparing Crash Report
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
}

IfWinExist Reporting a Crash
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
}

;OK
;GoodSync has crashed just now and we are sorry for that.
;Click OK to assemble Crash Report.
;You will be given option to submit it to Siber Systems.
;Submitting Crash Reports helps us in fixing these crashes.

IF (TRUE=FALSE)
{
	DetectHiddenWindows, OFF
	IfWinExist GoodSync
	{
		ControlGetText, OutputVar, GoodSync has crashed just now , GoodSync
		IF OutputVar 
		{
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			ControlClick, OK, GoodSync
			;WINHIDE
			
			;Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
		}
	}
}

DetectHiddenWindows, % dhw

Return

;--------------------------------------------------------------------
TIMER_SUB_WSCRIPT:
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON

IfWinExist Windows Script Host
{
	ControlClick, OK, Windows Script Host
	SoundBeep , 2500 , 100
}

DetectHiddenWindows, % dhw

Return


;--------------------------------------------------------------------
TIMER_SUB__MY_IP:

setTimer TIMER_SUB__MY_IP, % -1 * 1000 * 60 * 10 ; After10Minute

FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 23-MY IP.VBS"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%
	}

RETURN

;--------------------------------------------------------------------
TIMER_SUB__SendSMTP__0__LOG_BAT:

setTimer TIMER_SUB__SendSMTP__0__LOG_BAT, % -1 * 1000 * 60 * 60 ; 1 HOUR

FN_VAR:="C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP__0__LOG.BAT"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%, , MIN
	}

RETURN

;--------------------------------------------------------------------
TIMER_SUB_I_VIEW32_CONVERT_CCSE:

setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE, % -1 * 1000 * 60 * 30 ; HALF HOUR

if OSVER_N_VAR<10
{
	setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE,off
	RETURN
}


FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 24-I_VIEW32 CONVERT_CCSE.AHK"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%
	}

RETURN

;--------------------------------------------------------------------
TIMER_SUB_VICE_VERSA:

setTimer TIMER_SUB_VICE_VERSA, % -1 * 1000 * 60 * 60 ; After1Hours

VICE_VERSA_EXE=C:\Program Files\ViceVersa Pro\ViceVersa.exe
VICE_VERSA_PATH=C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\# VV C DRIVE ROOT

; Run, "%VICE_VERSA_EXE%" "%VICE_VERSA_PATH%\VV C DRIVE ROOT __ %A_ComputerName%.fsf" /hiddenautoexec

; /hiddenautoexec
; /dialogautoexec /autoclose 

Return

;----------------------------------------
TIMER_SUB_CMD_KILL:

;setTimer TIMER_SUB_CMD_KILL, % -1 * 1000 * 60 * 60 ; After1Hours

If (A_Now>START_CMD_KILL)
{
	IF (OSVER_N_VAR = 5)  ; WIN XP
	{
		START_CMD_KILL=%A_Now%
		START_CMD_KILL+=10*60 ; 10 MINUTES

		;Run, "TASKKILL.exe" /F /IM CMD.EXE /T , , HIDE
	}
}
Return


; -------------------------------------------------------------------
; Intel(R) Dynamic Platform and Thermal Framework Utility Application
; PROBLEM KEEP REPEAT RUNNING FOR SHORT SPACE FREQUENTLY AND REDUCE CPU COST
; Intel Corporation
; C:\Windows\Temp\DPTF\esif_assist_64.exe
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; STOP THE PROCESS WITH SUSPEND FOR AN HOUR AND LET BREATHE AGAIN 
; FOR A FEW MINUTE AND LOOPER
; ON MY OTHER COMPUTER esif_assist_64.exe EXISTS BUT DOESN'T KEEP GOING OFF ON
; [ Sunday 19:28:30 Pm_07 October 2018 ]
; -------------------------------------------------------------------
; -------------------------------------------------------------------
TIMER_SUB_ESIF_ASSIST_64_SUSPEND:
	Process, Exist, esif_assist_64.exe
	NewPID = %ErrorLevel%  ; Save the value immediately ErrorLevel is often changed

	;----------------------------------------------------------------
	; CHECK THE PID NUMBER & COMPARE IF IT NOT CHANGING THEN WE HAPPY 
	; IT NOT MOVING AN IN PROCESS SUSPEND SO WON'T REQUIRE ANY ACTION
	; SEEING AS THAT THE STATE IS ABLE TO BE FOUND TO BE IN
	;----------------------------------------------------------------
	If NewPID > 0 
		If OLDPID_ESIF_ASSIST_64_SUSPEND <> %NewPID%
			SET_GO_ESIF_ASSIST_64_SUSPEND+=1
	
	OLDPID_ESIF_ASSIST_64_SUSPEND = %NewPID%
	
	If SET_GO_ESIF_ASSIST_64_SUSPEND > 3
		If NewPID >0 
		{
			SET_GO_ESIF_ASSIST_64_SUSPEND=0
			;SoundBeep , 3000 , 100
			;SoundBeep , 3200 , 100
			Process_Suspend_PID(NewPID)
			SETTIMER TIMER_SUB_ESIF_ASSIST_64_SUSPEND_WAIT_AN_HOUR,OFF
			SETTIMER TIMER_SUB_ESIF_ASSIST_64_SUSPEND_WAIT_AN_HOUR,3600000 ; ---- 1 HOUR
		}
Return
;--------------------------------------------------------------------

;--------------------------------------------------------------------
TIMER_SUB_ESIF_ASSIST_64_SUSPEND_WAIT_AN_HOUR:
	;SoundBeep , 4000 , 100
	;SoundBeep , 4200 , 100
	Process_Resume("esif_assist_64.exe")
Return
;--------------------------------------------------------------------



; CREDIT DUE FIND
;--------------------------------------------------------------------
; Process, Suspend/Resume, example.exe - Suggestions - AutoHotkey Community
; https://autohotkey.com/board/topic/30341-process-suspendresume-exampleexe/
; --------
; Calculate Duration Between Two Dates – Results
; https://www.timeanddate.com/date/durationresult.html?d1=7&m1=10&y1=2018&d2=7&m2=10&y2=2018&h1=0&i1=0&s1=0&h2=&i2=1&s2=0
; --------
; 60*1000 - Google Search
; https://www.google.co.uk/search?ei=u0m6W-vGNMvSgAam84bYBg&q=60*1000&oq=60*1000&gs_l=psy-ab.3..0i7i30k1l4j0i7i10i30k1l2j0i7i30k1l3j0i30k1.122777.123404.0.124083.2.2.0.0.0.0.82.161.2.2.0....0...1c.1.64.psy-ab..0.2.159...0i8i30k1.0.xKF-kqsmqQw
;--------------------------------------------------------------------

;============================== Working on WinXP+
Process_Suspend(PID_or_Name){
    PID := (InStr(PID_or_Name,".")) ? ProcExist(PID_or_Name) : PID_or_Name
    h:=DllCall("OpenProcess", "uInt", 0x1F0FFF, "Int", 0, "Int", pid)
    If !h   
        Return -1
    DllCall("ntdll.dll\NtSuspendProcess", "Int", h)
    DllCall("CloseHandle", "Int", h)
}
Process_Resume(PID_or_Name){
    PID := (InStr(PID_or_Name,".")) ? ProcExist(PID_or_Name) : PID_or_Name
    h:=DllCall("OpenProcess", "uInt", 0x1F0FFF, "Int", 0, "Int", pid)
    If !h   
        Return -1
    DllCall("ntdll.dll\NtResumeProcess", "Int", h)
    DllCall("CloseHandle", "Int", h)
}

Process_Suspend_PID(PID){
    h:=DllCall("OpenProcess", "uInt", 0x1F0FFF, "Int", 0, "Int", pid)
    If !h   
        Return -1
    DllCall("ntdll.dll\NtSuspendProcess", "Int", h)
    DllCall("CloseHandle", "Int", h)
}
Process_Resume_PID(PID){
    h:=DllCall("OpenProcess", "uInt", 0x1F0FFF, "Int", 0, "Int", pid)
    If !h   
        Return -1
    DllCall("ntdll.dll\NtResumeProcess", "Int", h)
    DllCall("CloseHandle", "Int", h)
}

ProcExist(PID_or_Name=""){
    Process, Exist, % (PID_or_Name="") ? DllCall("GetCurrentProcessID") : PID_or_Name
    Return Errorlevel
}



;----------------------------------------
TIMER_SUB_SCRIPT_SHELL_FOLDERING:
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_SCRIPT_SHELL_FOLDERING,% -1 * 1000 * 60 * 60 ; After1Hours

; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 13-COPY MOVE SHELL FOLDING.VBS" , , hide

Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK" /RUN

; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK" , , hide
; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.VBS

Return


;----------------------------------------
TIMER_SUB_LOGGER:

setTimer TIMER_SUB_LOGGER, % -1 * 1000 * 60 * 60 ; After24Hours

Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 31-Duplicate Cleaner Logger.BAT" , , hide

Run, "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BlueToothView_Desc.VBS" , , hide

Return

;----------------------------------------
TIMER_SUB_WINDOWS_UPDATE:

;----
;Downloads / Tools by other people / Windows 10 Update Disabler
;https://winaero.com/download.php?view.1932
;----

RETURN

IF (A_ComputerName="7-ASUS-GL522VW") 
	RETURN

setTimer TIMER_SUB_WINDOWS_UPDATE, % -1 * 1000 * 60 * 60 ; After24Hours

OSVER_N_RESULT=0
IF (OSVER_N_VAR = 5)  ; WIN XP
	OSVER_N_RESULT=1
IF (OSVER_N_VAR = 6)  ; WIN 07
	OSVER_N_RESULT=1
IF (OSVER_N_VAR < 10) ; WIN 10
	OSVER_N_RESULT=1  ; SET NOT RUN IF THIS LESS THAN WINDOWS 10 OS SYS
	
COMP_N=0
;IF (A_ComputerName="7-ASUS-GL522VW") 
;	COMP_N=1        ; SET NOT RUN IF THIS COMPUTER

ALLOW_UPDATE_KILLER=1
IF (OSVER_N_RESULT=1) or (COMP_N=1) 
	ALLOW_UPDATE_KILLER=0

If (ALLOW_UPDATE_KILLER=1)
{
	Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 12-KEEP WINDOWS UPDATE STOPPED.BAT" /QUITE , , HIDE
}
Return


;----------------------------------------
TIMER_SUB_OWNER:

; IfWinExist TeamViewer Panel ahk_class TV_ControlWin
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
WinGet, ID, list,TeamViewer Panel ahk_class TV_ControlWin
DetectHiddenWindows, % dhw
IF ID_OLD_TeamViewer_Panel_TV_ControlWin<>%ID%
{
		ID_TeamViewer_Panel_TV_ControlWin_TIMER=%A_Now%
		; ID_TeamViewer_Panel_TV_ControlWin_TIMER+=10*60 ; 10 MINUTES
		ID_TeamViewer_Panel_TV_ControlWin_TIMER+=20 ; 20 SECOND MINUTES

		ID_OLD_TeamViewer_Panel_TV_ControlWin=%ID%
}
		
If (A_Now<ID_TeamViewer_Panel_TV_ControlWin_TIMER)
{
	Process, Exist, ICACLS.EXE
	If ErrorLevel > 0
	{
		Process, Close, ICACLS.EXE
		SoundBeep , 2000 , 100
		ID_TeamViewer_Panel_TV_ControlWin_TIMER=%A_Now%
		ID_TeamViewer_Panel_TV_ControlWin_TIMER+=20 ; 20 SECOND MINUTES
	}
}

IF (OSVER_N_VAR < 6 ) ; THAN XP
	RETURN
		
IF TIMER_SUB_OWNER_SAVE_TIMER<%A_NOW%
{	
	TIMER_SUB_OWNER_SAVE_TIMER=%A_NOW%
	TIMER_SUB_OWNER_SAVE_TIMER+= 2, Days
	IF (A_ComputerName="3-LINDA-PC") 
		TIMER_SUB_OWNER_SAVE_TIMER+= 4, Days

	FileDELETE, %SCRIPT_NAME_VAR%
	FileAppend,%TIMER_SUB_OWNER_SAVE_TIMER%,%SCRIPT_NAME_VAR%

}
ELSE
{
	RETURN
}
	
SET_GO_RESULT=0

IF (OSVER_N_VAR > 5 ) ; BIGGER THAN XP
	SET_GO_RESULT=1

Process, Exist, ICACLS.EXE
If ErrorLevel > 0
	SET_GO_RESULT=0

If (SET_GO_RESULT=1)
{
	Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\#_OWNER-HARDCODED ANYWHERE.BAT" /QUITE , , HIDE
}

RETURN



TIMER_Check_Any_PID_Suspended_Warning:
	SETTIMER TIMER_Check_Any_PID_Suspended_Warning, 3600000 ; ---- 10 SECONDS ---- And Then 1 Hour

	Element_1=C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk

	SET_GO=FALSE
	IfExist, %Element_1%
		IF !WinExist(Element_1) 
			SET_GO=TRUE

	IF SET_GO=TRUE	
		{
			Run, "%Element_1%" /QUITE_COMMANDLINE_ARGS
		}
RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; END OF CODE ONLY THE END BLOCK ROUTINE FOR EXIT APP HERE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,59000

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

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

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
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------


; REFERENCE SET
; -------------------------------------------------------------------
; --------------------------------------------------------------
;GOOD SCRIPT EXAMPLE PAGE HELPER
;----
;1 Hour Software by Skrommel - DonationCoder.com
;http://www.donationcoder.com/Software/Skrommel/index.html#GoneIn60s
;----

;StringMid,now,A_Now,9,4
; -------------------------------------------------------------------
; --------------------------------------------------------------
; [ Sunday 19:26:10 Pm_15 April 2018 ]
; THIS IS OUR PREFERRED DEFAULT OPTIONS FOR INSTALLING NOTEPAD++
; THE MIDDLE CHECKBOX IS SELECTED
; NONE OF OUR PLUGIN WILL WORK PROPER AS OUR SETUP USES THE
; Allow plugins to be loaded from %APPDATA%\notepad++ -- CHECKBOX CHECKED
; EVERY-TIME AN UPDATE COMES ALONG HAS TO BE REMEMBER
; OUR SETTER USES 32-BIT NOTEPAD FOR SOME REASON NOT UPGRADED YET
; ----
; AUTOHOTKEYS SELECT A CHECKBOX - Google Search
; https://www.google.co.uk/search?q=AUTOHOTKEYS+SELECT+A+CHECKBOX&oq=AUTOHOTKEYS+SELECT+A+CHECKBOX&aqs=chrome..69i57j0.11108j0j7&sourceid=chrome&ie=UTF-8
; --------
; Check a Checkbox - Control check..... - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/25075-check-a-checkbox-control-check/
; ----
; --------------------------------------------------------------
