;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 75-GOODSYNC OPTIONS SET.ahk
;# __ 
;# __ Autokey -- 75-GOODSYNC OPTIONS SET.ahk
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
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
; -------------------------------------------------------------------
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
;
; FROM __ Sun 28-Apr-2019 08:40:38
; TO   __ Sun 28-Apr-2019 08:40:38 __ 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 002 AND 003 
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

; IF A_ComputerName=2-ASUS-EEE
	; Exitapp

GLOBAL OSVER_N_VAR

GOODSYNC_HANDLE_CHECK_CHANGE_OLD_ONE=0
GOODSYNC_HANDLE_CHECK_CHANGE_OLD_2GO=0
O_Status_GSDATA=0
O_HWND_1=0

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6


; IF A_ComputerName=1-ASUS-X5DIJ
; IF A_ComputerName=2-ASUS-EEE


; GOODSYNC OPTIONS
; RECOVERY FROM DAMAGE DONE BY CHANGE USER NAME FOR GOODSYNC DIRECT SERVER 
; ------------------------------------------------------------------------
; SETTIMER SET_OK_BOX,800


; -------------------------------------------------------------------
SETTIMER TIMER_SUB_GOODSYNC_OPTIONS,1000
SETTIMER TIMER_SUB_GOODSYNC,1000
SETTIMER MINIMIZE_AND_RUN_GOODSYNC_V10,10000
SETTIMER MINIMIZE_AND_RUN_GOODSYNC_2GO,10000


RETURN

; SET OKAY BOX AFTER MADE SELECTION
SET_OK_BOX:
{

	DetectHiddenWindows, ON
	SetTitleMatchMode 3

	WinGet, HWND_1, ID, A
	
	IfWinExist Left Folder ahk_class #32770
	WinGet, HWND_2, ID, Left Folder ahk_class #32770
	IF HWND_2>0 
	IF HWND_1<>%HWND_2%
		WinActivate,  ahk_id %HWND_2%
	IfWinExist Right Folder ahk_class #32770
	WinGet, HWND_2, ID, Right Folder ahk_class #32770
	IF HWND_2>0 
	IF HWND_1<>%HWND_2%
		WinActivate,  ahk_id %HWND_2%
	SetTitleMatchMode 2
	IfWinExist ] Options ahk_class #32770
	WinGet, HWND_2, ID, ] Options ahk_class #32770
	IF HWND_2>0 
	IF HWND_1<>%HWND_2%
		WinActivate,  ahk_id %HWND_2%
	
	IF HWND_2=0 
		RETURN

	IfWinExist Left Folder ahk_class #32770
		WinActivate, Left Folder ahk_class #32770
	IfWinActive Left Folder ahk_class #32770
	{	
		; ControlGetPos, x, y, , , OK, Left Folder ahk_class #32770
		; MouseMove, X+10, Y+10
		ControlClick, OK, Left Folder ahk_class #32770,,,, NA x20 y20
		SoundBeep , 2000 , 400
	}	
	
	IfWinExist Right Folder ahk_class #32770
		WinActivate, Right Folder ahk_class #32770
	IfWinActive Right Folder ahk_class #32770
	{	
		; ControlGetPos, x, y, , , OK, Right Folder ahk_class #32770
		; MouseMove, X+10, Y+10
		ControlClick, OK, Right Folder ahk_class #32770,,,, NA x20 y20
		SoundBeep , 3000 , 400
	}	

	
	
	DetectHiddenWindows, ON
	SetTitleMatchMode 2
	IfWinExist Options ahk_class #32770
	WinActivate, Options ahk_class #32770

	IfWinActive ] Options ahk_class #32770
	{	
		ControlGettext, OutputVar_2, Button17, ] Options ahk_class #32770
		ControlGet, OutputVar_1, Line, 1, Edit9, ] Options ahk_class #32770
		; -----------------------------------------------------------
		Periodic_TIMER_VALUE=5
		; -----------------------------------------------------------
		IF (OutputVar_2="Periodically (On Timer), every")
		IF OutputVar_1<>%Periodic_TIMER_VALUE%
		{
			ControlSetText, Edit9,, ] Options ahk_class #32770
			Control, EditPaste, %Periodic_TIMER_VALUE%, Edit9, ] Options ahk_class #32770
			SoundBeep , 4000 , 100
		}
		
		; ClassNN:	Button17
		; Text:	Periodically (On Timer), every
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			ControlGet, Status, Checked,, Button17, ] Options ahk_class #32770
			If Status=0
			{
				Control, Check,, Button17, ] Options ahk_class #32770
				SoundBeep , 4000 , 100
			}
		}

		; MSGBOX % OutputVar_2

		IF OutputVar_2
		IF (OutputVar_2<>"Periodically (On Timer), every")
			MSGBOX Button17 NOT ANY LONGER ASSOCIATED WITH`nPeriodically (On Timer), every

		
		If Status=1
		IF OutputVar_1=%Periodic_TIMER_VALUE%
		IfWinActive ] Options ahk_class #32770
		{	
			OutputVar_2=
			ControlGettext, OutputVar_2, Button61, ] Options ahk_class #32770
			IF (OutputVar_2="Save")
				ControlClick, Button61, Options ahk_class #32770,,,, NA x10 y10
				
			OutputVar_2=
			ControlGettext, OutputVar_2, Button62, ] Options ahk_class #32770
			IF (OutputVar_2="Save")
				ControlClick, Button62, Options ahk_class #32770,,,, NA x10 y10
			
			OutputVar_2=
			ControlGettext, OutputVar_2, Button63, ] Options ahk_class #32770
			IF (OutputVar_2="Save")
				ControlClick, Button63, Options ahk_class #32770,,,, NA x10 y10
			
			SoundBeep , 5000 , 400
		}	
	}
		
	
}
RETURN


; -------------------------------------------------------------------
TIMER_SUB_GOODSYNC_OPTIONS:


dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.


WinGet, HWND_1, ID, ] Options ahk_class #32770
IF HWND_1>0
{
	WinGet, HWND_1_EXENAME, ProcessName, ahk_id %HWND_1%
	WinGet, HWND_2, ID, A
	IF HWND_2=%HWND_1%
	{
		; WinGet, OutputVar, ControlList, ahk_id %HWND_1%
		; Tooltip, % OutputVar ; List All Controls of Active Window
		;---------------------------------------------------------
		ControlGettext, OutputVar_2, Button17, ahk_id %HWND_1%

		ControlGet, OutputVar_1, Line, 1, Edit9, ahk_id %HWND_1%
		
		WinGetTitle OutputVar_3,ahk_id %HWND_1%
		
		HWND_1_EXENAME_GoodSync_v10_exe_DONE=FALSE
		SET_GO=FALSE
		IF A_ComputerName=8-MSI-GP62M-7RD
		IF HWND_1_EXENAME=GoodSync-v10.exe
		{
			IF OutputVar_1=2
				SET_GO=TRUE
			IF OutputVar_1=5
				SET_GO=TRUE
		}
		
		IF OutputVar_2
		IF (OutputVar_2<>"Periodically (On Timer), every")
			MSGBOX Button17 NOT ANY LONGER ASSOCIATED WITH`nPeriodically (On Timer), every

		IF SET_GO=TRUE
	    IF (OutputVar_2="Periodically (On Timer), every")
		{
			HWND_1_EXENAME_GoodSync_v10_exe_DONE=TRUE
			ControlSetText, Edit9,, ahk_id %HWND_1%
			Control, EditPaste, 4, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100

		}

		IF HWND_1_EXENAME_GoodSync_v10_exe_DONE=FALSE
		IF HWND_1_EXENAME=GoodSync-v10.exe
		IF OutputVar_1=2
	    IF (OutputVar_2="Periodically (On Timer), every")
		{
			ControlSetText, Edit9,, ahk_id %HWND_1%
			Control, EditPaste, 5, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100

		}
		SET_GO=FALSE
		IF HWND_1_EXENAME=GoodSync2Go.exe
		{
			IF OutputVar_1=1
				SET_GO=TRUE
			IF OutputVar_1=2
				SET_GO=TRUE
		}
		IF SET_GO=TRUE
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			ControlSetText, Edit9,, ahk_id %HWND_1%
			Control, EditPaste, 5, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
		
			
		; ; PRESS SAVE WHEN SETTING OPTIONS DONE
		; ControlGet, OutputVar_1, Line, 1, Edit9, ahk_id %HWND_1%
		; ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
		; If (OutputVar_1 = 5 and Status=1)
		; {
			; ControlGetPos, x, y, , , Button65, ahk_id %HWND_1%
			; MouseMove, X+10, Y+10		
			; ControlClick, Button65,ahk_id %HWND_1% ; SAVE 
			; SoundBeep , 4000 , 100
		; }	
		
		; WinGet, HWND_1, ID, ] Options ahk_class #32770
		; TOOLTIP % OutputVar_3
		If INSTR(OutputVar_3,"HDD HUBIC")>0 
		{
			ControlGet, Status, Checked,, Button58, ahk_id %HWND_1%
			If Status=1
			{
				Control, UNCheck,, Button58, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}

			; ControlGet, Status, Checked,, Button58, ahk_id %HWND_1%
			; If Status=0
			; {
				; LOOP
				; {
					; SLEEP 50
					; ControlGetPos, x, y, , , Button65, ahk_id %HWND_1%
					; MouseMove, X+10, Y+10
					
					; ControlClick, Button65,ahk_id %HWND_1%
					; IfWinNotExist, ahk_id %HWND_1%
						; BREAK
					; SoundBeep , 4000 , 50
				; }
			; }
		}
		
		
		; O_HWND_1=0
		if O_HWND_1<>%HWND_1%
		{
			; ClassNN:	Button17
			; Text:	Periodically (On Timer), every

			IF OutputVar_2
			IF (OutputVar_2<>"Periodically (On Timer), every")
				MSGBOX Button17 NOT ANY LONGER ASSOCIATED WITH`nPeriodically (On Timer), every
			
			IF (OutputVar_2="Periodically (On Timer), every")
			{
				ControlGet, OutputVar_4, Visible, , Button17, ahk_id %HWND_1%
				ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
				If Status=0
				{
					Control, Check,, Button17, ahk_id %HWND_1%
					IF OutputVar_4=1
						SoundBeep , 4000 , 100
				}
			}
		}

		; {
			; ControlGet, OutputVar_4, Visible, , Button17, ahk_id %HWND_1%
			; ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
			; If Status=0
			; {
				; Control, Check,, Button17, ahk_id %HWND_1%
				; IF OutputVar_4=1
					; SoundBeep , 4000 , 100
			; }
		; }
		
		ControlGettext, OutputVar_2,F Button21, ahk_id %HWND_1%
		ControlGet, OutputVar_1, Line, 1, Edit11, ahk_id %HWND_1%
		
		If (OutputVar_1 = 100 ; DEFAULT
			and OutputVar_2="Do not Sync if changed files more than")
			{
				ControlSetText, Edit11,, ahk_id %HWND_1%
				Control, EditPaste, 80,	Edit11, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
		}
		; ControlGet, Status, Checked,, Button21, ahk_id %HWND_1%
		; If Status=0
		; {
			; Control, Check,, Button21, ahk_id %HWND_1%
				; SoundBeep , 4000 , 100
		; }
		
		; ; PRESS SAVE WHEN SETTING OPTIONS DONE
		; ControlGet, OutputVar_1, Line, 1, Edit11, ahk_id %HWND_1%
		; ControlGet, Status, Checked,, Button21, ahk_id %HWND_1%
		; If (OutputVar_1 = 90 and Status=1)
		; {
			; ControlGetPos, x, y, , , Button65, ahk_id %HWND_1%
			; MouseMove, X+10, Y+10		
			; ControlClick, Button65,ahk_id %HWND_1% ; SAVE 
			; SoundBeep , 4000 , 100
		; }
		
		
		ControlGettext, OutputVar_2, Button23, ahk_id %HWND_1%
		ControlGet, OutputVar_1, Line, 1, Edit12, ahk_id %HWND_1%

		IF OutputVar_2
		IF (OutputVar_2<>"Wait for Locks to clear, minutes")
			MSGBOX Button23 NOT ANY LONGER ASSOCIATED WITH`nWait for Locks to clear, minutes
		
		
		If (OutputVar_1<>20 
			and OutputVar_2="Wait for Locks to clear, minutes")
			{
				ControlSetText, Edit12,, ahk_id %HWND_1%
				Control, EditPaste, 20, Edit12, ahk_id %HWND_1%
				SoundBeep , 4000 , 100

		}
		ControlGet, Status, Checked,, Button23, ahk_id %HWND_1%
		If Status=0
		{
			Control, Check,, Button23, ahk_id %HWND_1%
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
			ControlGet, Status, Checked,, *4, ahk_id %HWND_1%




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
}

;Are you sure you want to move _GSDATA_ folder back to the sync folder?
;This option is for Advanced users and you should read the manual first.
;Yes
;No

;WinGet, HWND_1, ID, GoodSync Warning ahk_class #32770
WinGet, HWND_1, ID, GoodSync ahk_class #32770
IF HWND_1>0 
{
	WinGet, HWND_2, ID, A
	IF HWND_2=%HWND_1%
	{
		WinGetText OutputVar_3,ahk_id %HWND_1%

		IfInString, OutputVar_3, Are you sure you want to move _GSDATA_
		{
			ControlClick, Button2,ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}

		IfInString, OutputVar_3, Are you sure you want to keep _GSDATA_
		{
			ControlClick, Button2,ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}

		WinGet, HWND_1, ID, GoodSync ahk_class #32770
		WinGetText OutputVar_3,ahk_id %HWND_1%

		IfInString, OutputVar_3, Removable drive with volume name
		{
			ControlClick, Yes,ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}


		; We recommend not to sync to disk root folder, because:
		; - there are limitations on how many files and folders you can have in disk root folder,
		; - sync folder name would help you identify copy of what folder you keep in there.
		; So we will sync to folder K:\=D_DRIVE-2TB - Backup instead. Is this Ok?
		; Yes
		; No
		IfInString, OutputVar_3, We recommend not to sync to disk root folder
		{
			ControlClick, Button3,ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
	}
}
DetectHiddenWindows, % dhw
Return


MINIMIZE_AND_RUN_GOODSYNC_V10:

; GET GOODSYNC HANDLE AND THEN CHECK IF CHANGE DUE TO UPDATE HAPPEN
; IF HAS AND THEN MINIMIZE ON CERTAIN COMPUTER
; -------------------------------------------------------------------
DetectHiddenWindows, ON
SET_GO=TRUE
	
IF SET_GO=TRUE
{
	WinGet, HWND_1, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}

	IF HWND_1>0 
	{
		If GOODSYNC_HANDLE_CHECK_CHANGE_OLD_ONE <> %HWND_1%
		{
			WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
			; IF A_ComputerName=2-ASUS-EEE
				; If MMX<>-1
					; MSGBOX % MMX 
			
			If MMX<>-1
			{
				; -------------------------------------------------------------
				; IF A_ComputerName=2-ASUS-EEE
				; -------------------------------------------------------------
				; OF HERE LOW THE END EEE ON WIN-XP IT DOESN'T TAKE THE REQUEST 
				; TO MINIMIZE AS ONE COMMAND
				; AND HAS TO REPEAT UNTIL DONE IT
				; EXAMPLE LEFT WITH WINDOW UP IN MAXIMUM AND NOT ISSUE IT COMMAND TO DO MINIMIZED
				; UNTIL REPEAT UNTIL FEW TIME AND THEN STOP
				; [ Thursday 14:05:20 Pm_14 March 2019 ]
				; -------------------------------------------------------------
				WinMinimize  ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
				SOUNDBEEP 2000,100
				SLEEP 2000

				WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
				If MMX<>-1
				{
					HWND_1=0
				}
			}
		}

		; ------------------------------------------------------------------------
		; SO SOMETHING TURNED UP I PUT AND _IF_ IN VARIABLE NAME AND WOULDN'T WORK
		; ------------------------------------------------------------------------
		
		GOODSYNC_HANDLE_CHECK_CHANGE_OLD_ONE = %HWND_1%
	}
}

DetectHiddenWindows, ON

IF (TRUE=TRUE)
{
	; ---------------------------------------------------------------
	; PROBLEM ONE COMPUTER MY MSI INTEL 7 WIN 10 
	; HAS THAT PROCESS IS CALLED GoodSync.exe
	; WHILE OTHER COMPUTER USE   GoodSync-v10.exe
	; SO WE __ 19 JAN 2019 13:50
	; ---------------------------------------------------------------
	SET_GO=TRUE
	Process, Exist, GoodSync-v10.exe
	If ErrorLevel
		SET_GO=FALSE
	Process, Exist, GoodSync.exe
	If ErrorLevel
		SET_GO=FALSE

	IFWINEXIST ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
		SET_GO=FALSE
	
	IF SET_GO = TRUE
	{
		FN_VAR:="C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			; MSGBOX HERE
			Run, "%FN_VAR%" , , MIN
		}
	}
}

RETURN


MINIMIZE_AND_RUN_GOODSYNC_2GO:

; GET GOODSYNC HANDLE AND THEN CHECK IF CHANGE DUE TO UPDATE HAPPEN
; IF HAS AND THEN MINIMIZE ON CERTAIN COMPUTER
; -------------------------------------------------------------------
DetectHiddenWindows, ON
IF (A_ComputerName<>"7-ASUS-GL522VW" or A_UserName="MATT 04")
	RETURN
	
WinGet, HWND_1, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}

IF HWND_1>0 
{
	If GOODSYNC_HANDLE_CHECK_CHANGE_OLD_2GO <> %HWND_1%
	{
		WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
		; IF A_ComputerName=2-ASUS-EEE
			; If MMX<>-1
				; MSGBOX % MMX 
		
		If MMX<>-1
		{
			; -------------------------------------------------------------
			; IF A_ComputerName=2-ASUS-EEE
			; -------------------------------------------------------------
			; OF HERE LOW THE END EEE ON WIN-XP IT DOESN'T TAKE THE REQUEST 
			; TO MINIMIZE AS ONE COMMAND
			; AND HAS TO REPEAT UNTIL DONE IT
			; EXAMPLE LEFT WITH WINDOW UP IN MAXIMUM AND NOT ISSUE IT COMMAND TO DO MINIMIZED
			; UNTIL REPEAT UNTIL FEW TIME AND THEN STOP
			; [ Thursday 14:05:20 Pm_14 March 2019 ]
			; -------------------------------------------------------------
			WinMinimize  ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
			SOUNDBEEP 2000,100
			SLEEP 2000

			WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
			If MMX<>-1
			{
				HWND_1=0
			}
		}
	}

	; ------------------------------------------------------------------------
	; SO SOMETHING TURNED UP I PUT AND _IF_ IN VARIABLE NAME AND WOULDN'T WORK
	; ------------------------------------------------------------------------
	GOODSYNC_HANDLE_CHECK_CHANGE_OLD_2GO = %HWND_1%
}

DetectHiddenWindows, ON

; ---------------------------------------------------------------
; PROBLEM ONE COMPUTER MY MSI INTEL 7 WIN 10 
; HAS THAT PROCESS IS CALLED GoodSync.exe
; WHILE OTHER COMPUTER USE   GoodSync-v10.exe
; SO WE __ 19 JAN 2019 13:50
; ---------------------------------------------------------------
SET_GO=TRUE
Process, Exist, GoodSync2Go.exe
If ErrorLevel
	SET_GO=FALSE

IFWINEXIST ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
	SET_GO=FALSE

IF SET_GO = TRUE
{
	FN_VAR:="C:\GoodSync\x64\GoodSync2Go.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 4000 , 100
		SoundBeep , 3000 , 100
		SoundBeep , 4000 , 100
		SoundBeep , 3000 , 100
		; MSGBOX HERE
		Run, "%FN_VAR%" , , MIN
	}
}

RETURN




;--------------------------------------------------------------------
TIMER_SUB_GOODSYNC:
;--------------------------------------------------------------------
; setTimer TIMER_SUB_GOODSYNC, OFF
dhw := A_DetectHiddenWindows
DetectHiddenWindows, OFF
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.





; -------------------------------------------------------------------
; GoodSync Script Command to Stop in Wait using a Messenger Box
; TO MAX AND THEN RESTORE WINDOW OF GOODSYNC WHEN GET TO END &
; PRESS BUTTON ON MESSAGE BOX THAT INDICATED END TO ALLOW CONTINUE AGAIN
; TRY AND GET OVER PROBLEM AND WINDOW WON'T REFRESH AFTER RUN BIG 
; LONG JOB AFTER ABOUT ONE DAY
; IF ANYTHING ELSE FAIL THEN RESTART GOODSYNC WHEN END OF WORK
; RESTART WILL JUST 
; REQUIRE CLOSE AND LINE SET ABOVE WILL DO 
; THAT RESTART THING
; -------------------------------------------------------------------
; GoodSync Script Command to Stop in Wait using a Messenger Box

DetectHiddenWindows, OFF
SetTitleMatchMode 2  

VAR_WORKER_MSGBOX_DELAY_COUNT_02=ahk_class #32770 ahk_exe WScript.exe

OutputVar=
IF (A_ComputerName="7-ASUS-GL522VW") 
{
	IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
	{
		ControlGetText, OutputVar, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
		ControlGettext, OutputVar_2, Button1, %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
	}
	IF Instr(OutputVar,"GoodSync Script Command to Stop")
	IF INSTR(OutputVar_2,"&Yes  0")
	{
		WinMaximize, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
		SLEEP 4000
		WinRESTORE, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
		SLEEP 4000
	
		; CLICK THE MESSENGER BOX
		ControlClick, &Yes  0, %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
		SOUNDBEEP 1500,50
		
		; CLOSE GOODSYNC
		ControlGetText, OutputVar, %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
		IF Instr(OutputVar,"GoodSync Script Command to Stop")=0
		{
			SOUNDBEEP 1500,50
			WinClose, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
		}
	}
}
 	
OutputVar=
IF (A_ComputerName="7-ASUS-GL522VW") 
	IFWINEXIST GoodSync ahk_class #32770
	{
		ControlGetText, OutputVar, Edit1, GoodSync ahk_class #32770
		IF Instr(OutputVar,"One or more jobs are running now")
		{
			SoundBeep , 2000 , 100
			SoundBeep , 1500 , 100
			ControlClick, Button2 , ahk_class #32770 ; OK
				
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
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			If ( PID_8 <> DllCall("GetCurrentProcessId") ) 
				 ; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				 Process, Close, %PID_8% 
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

if A_ThisMenuItem=Pause __ Debby Hall
{
	IF DEBBY_HALL_PAUSE=TRUE
	{
		DEBBY_HALL_PAUSE=FALSE
		SOUNDBEEP 5000,200
	}
	ELSE
	{
		DEBBY_HALL_PAUSE=TRUE
		SOUNDBEEP 1000,200
	}
}
return


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

