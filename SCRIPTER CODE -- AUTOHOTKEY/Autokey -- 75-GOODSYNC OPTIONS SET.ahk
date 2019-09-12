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

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 004
; -------------------------------------------------------------------
; NEW SUBROUTINE SET WORK ON TODAY
; -------------------------------------------------------------------
; GOSUB SET_ACTIVATION_BOX_GOODSYNC_TEXT_FILE
; SETTIMER TIMER_SET_ACTIVATION_BOX,1000
; -------------------------------------------------------------------
; THIS WILL PUT ACTIVATION KEY ON GOODSYNC
; WHICH IS HANDY AS GOODSYNC2GO ON D DRIVE WHEN SHARING MANY COMPUTER
; DOES APPEAR FREQUENTLY
; AND GOT 10 COPY OF GOODSYNC ABOUT 3 ARE GOODSYNC2GO 
; THE EXPENSIVE VERSION
; DOES GET HARD FIDDLY PUT ALL THE RIGHT SERIAL NUMBER IN
;
; THE CODE IF FILE BASE FOR KEY
; AND COMPLETE ARRAY BASE FOR EASIER 
; ALSO CHECK THE TEXT FILE FOR UPDATE TO INCLUDE CHANGER
;
; A MORNING WORK FOR ONE CHECK BOX IN
;
; GOODSYNC2GO IS LITTLE ANNOYER
; AS FOR PORTABLE
; WOULD THINK
; THAT SERIAL KEY NOT REQUIRE AS MUCH
; IF STOP DOWN ALL OTHER INSTANCE NETWORK
; SHOULD BE THAT NOT REQUIRE PUT KEY SO OFTEN
; MAYBE IS CHANGE COMPUTER NORMALLY USE ON
; SO IF OUT PORTABLE HAVE TO GET OUT KEY EVERYWHERE
; AND AUTO ENTRY NOT WORK
; -------------------------------------------------------------------
; Wed 04-Sep-2019 12:58:00
; -------------------------------------------------------------------
; FROM __ Wed 04-Sep-2019 10:34:52
; TO   __ Wed 04-Sep-2019 12:13:10
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 005
; -------------------------------------------------------------------
; TIDY UP OPTION SET FOR GOODSYNC
; -------------------------------------------------------------------
; ADD THIS ROUTINE
; -------------------------------------------------------------------
; EDITOR -- QUICK MADE NOT OPERATION AND TRANSFER TO 
; FOR AN ALL IN ONE ROUTINE WITH ARRAY PUT
; -------------------------------------------------------------------
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; TIMER_RENAME_EXTENSION_OF_MP4_TO_UPPER:
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; BIT SEMI OFF TOPIC FOR HERE
; JUST PLANT ONE IN THAN OWN CODE MAYBE ON TO IT
; STILL GOT TO DO WITH GOODSYNC FOR VB
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; HERE IS SIMILAR TO PROJECT DONE AFTER AND BETTER
; QUICK TURN INTO NOT OPERATION AND MOVE LOCATE HERE
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; TIMER_RENAME_EXTENSION_OF_MP4_TO_UPPER:
; LIKE HERE LINK BOTH IN CASE MORE PROGRESS TO BE DONE
; -------------------------------------------------------------------
; Autokey -- 75-GOODSYNC OPTIONS SET.ahk
; DETECT_VB_IDE_BEEN_RUN_MULTIPLE_AND_DO_TASK_AFTER:
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM __ Tue 10-Sep-2019 20:00:00
; TO   __ Tue 10-Sep-2019 20:15:00
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


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

; IF A_ComputerName=2-ASUS-EEE
	; Exitapp

GLOBAL TOOLTIP_SET_REMOVE_TIMER
TOOLTIP_SET_REMOVE_TIMER=	
	
GLOBAL OSVER_N_VAR

GLOBAL GOODSYNC_CHECK_CHANGE_OLD_ONE_HANDLE
GLOBAL GOODSYNC_CHECK_CHANGE_OLD_2GO_HANDLE
GLOBAL GOODSYNC_CHECK_CHANGE_OLD_2GO_PID
GLOBAL GOODSYNC_CHECK_CHANGE_OLD_ONE_PID
GLOBAL O_Status_GSDATA
GLOBAL O_HWND_1

TEMPORARY_FILE_HWND=
PERIODICALLY_SET_VALUE_HWND=
UNCHECK_PERIODIC_TIMER_HWND=
CHECK_ESTIMATE_DISK_SPACE_REQUIRED_HWND=
HDD_HUBIC_HWND=
DO_NOT_SYNC_IF_CHANGED_FILES_MORE_THAN_HWND=
WAIT_FOR_LOCKS_TO_CLEAR_MINUTE_HWND=

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

GLOBAL TIMER_NOT_RESPONDING
TIMER_NOT_RESPONDING=0

SET_ACTIVATION__HWND=0
ACTIVATION_GS_FILENAME_CONTENT=
O_WNDCLASS_DESKED_GSK_MATCH=


; -------------------------------------------------------------------
SETTIMER TIMER_SUB_GOODSYNC_OPTIONS,1000
SETTIMER TIMER_SUB_GOODSYNC_SCRIPT_COMMAND_TO_STOP,1000
SETTIMER MINIMIZE_AND_RUN_GOODSYNC_V10,10000
SETTIMER MINIMIZE_AND_RUN_GOODSYNC_2GO,10000
SETTIMER CHECK_GOODSYNC_NOT_RESPOND,2000
FILE_NAME_DATE_MOD_GOODSYNC_KEY_SCRIPT=0
GOSUB SET_ACTIVATION_BOX_GOODSYNC_TEXT_FILE
SETTIMER TIMER_SET_ACTIVATION_BOX,1000
SET_GOODSYNC_CONNECT_BOX_HWND=0
SETTIMER TIMER_SET_GOODSYNC_CONNECT_BOX,1000
SETTIMER TOOLTIP_REMOVER,OFF

; -------------------------------------------------------------------
RETURN
; -------------------------------------------------------------------

DETECT_VB_IDE_BEEN_RUN_MULTIPLE_AND_DO_TASK_AFTER:
RETURN
; -------------------------------------------------------------------
; HERE IS SIMILAR TO PROJECT DONE AFTER AND BETTER
; -------------------------------------------------------------------
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; TIMER_RENAME_EXTENSION_OF_MP4_TO_UPPER:
; LIKE HERE LINK BOTH IN CASE MORE PROGRESS TO BE DONE
; -------------------------------------------------------------------
; Autokey -- 75-GOODSYNC OPTIONS SET.ahk
; DETECT_VB_IDE_BEEN_RUN_MULTIPLE_AND_DO_TASK_AFTER:
; -------------------------------------------------------------------

DetectHiddenWindows, On
WinGet, WNDCLASS_DESKED_GSK_MATCH, List, ahk_class wndclass_desked_gsk

IF O_WNDCLASS_DESKED_GSK_MATCH
IF WNDCLASS_DESKED_GSK_MATCH<%O_WNDCLASS_DESKED_GSK_MATCH%
{
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 82-RENAME VB PROJECT VBP LOWER CASE.ahk"
	IfExist, %FN_VAR%
	{
		Run, "%FN_VAR%" , , MIN
	}
}
O_WNDCLASS_DESKED_GSK_MATCH=%WNDCLASS_DESKED_GSK_MATCH%
RETURN



TIMER_SET_GOODSYNC_CONNECT_BOX:
	DetectHiddenWindows, ON
	SetTitleMatchMode 3
	IfWinExist GoodSync Account Setup ahk_class #32770
	WinGet, HWND_GOODSYNC_CONNECT_BOX, ID, GoodSync Account Setup ahk_class #32770
	IF SET_GOODSYNC_CONNECT_BOX_HWND<>%HWND_GOODSYNC_CONNECT_BOX%
	IF HWND_GOODSYNC_CONNECT_BOX
	{
		Element_2=matt-lan-btinternet-com
		Element_3=24682468
		ControlGettext, OutputVar_2, Edit2, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
		IF OutputVar_2<>%Element_2%
		{
			ControlSetText, Edit2,, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
			Control, EditPaste, %Element_2%, Edit2, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
			SoundBeep , 4000 , 100
			Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
		ControlGettext, OutputVar_2, Edit3, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
		IF OutputVar_2<>%Element_3%
		{
			ControlSetText, Edit3,, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
			Control, EditPaste, %Element_3%, Edit3, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
			SoundBeep , 4000 , 100
			Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	
	IF SET_GOODSYNC_CONNECT_BOX_HWND<>%HWND_GOODSYNC_CONNECT_BOX%
	{
		ALL_SET_OF_GOODSYNC_CONNECT=TRUE

		ControlGettext, OutputVar_2, Edit2, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
		IF OutputVar_2<>%Element_2%
			ALL_SET_OF_GOODSYNC_CONNECT=FALSE
		ControlGettext, OutputVar_2, Edit3, ahk_id %HWND_GOODSYNC_CONNECT_BOX%
		; RESULT IS NONE SEEM A PASS BOX
		; NONE WAY LEARN IS COMPLETER		
		; -----------------------------------------------------------
		; TOOLTIP % OutputVar_2
		; IF OutputVar_2<>%Element_3%
			; ALL_SET_OF_GOODSYNC_CONNECT=FALSE
		IF ALL_SET_OF_GOODSYNC_CONNECT=TRUE
			SET_GOODSYNC_CONNECT_BOX_HWND=%HWND_GOODSYNC_CONNECT_BOX%
	}
RETURN


; -------------------------------------------------------------------
SET_ACTIVATION_BOX_GOODSYNC_TEXT_FILE:

	FN_Array_1 := []
	ArrayCount := 1
	File_NAME := "C:\RF\7-ASUS-GL522VW\SAFE NOTE\Autokey -- 75-GOODSYNC OPTIONS SET_GOODSYNC KEY.txt"
	IfNOTExist, %File_NAME%
		RETURN 
	
	FileRead, LoadedText, %File_NAME%
	oText := StrSplit(LoadedText, "`n")
	Loop, % oText.MaxIndex()
	{
		IF SUBSTR(oText[A_Index],1,1)<>"#"
		IF TRIM(oText[A_Index])<>""
		{
			STRING_V:=oText[A_Index]
			StringReplace, STRING_V, STRING_V, `n, , All
			StringReplace, STRING_V, STRING_V, `r, , All
			FN_Array_1[ArrayCount]:=STRING_V
			ArrayCount += 1
		}
	}
	FileGetTime, OutputVar, %File_NAME%, M
	IF FILE_NAME_DATE_MOD_GOODSYNC_KEY_SCRIPT<>%OutputVar%
		FILE_NAME_DATE_MOD_GOODSYNC_KEY_SCRIPT = %OutputVar%

	
RETURN

TIMER_SET_ACTIVATION_BOX:
{

	File_NAME := "C:\RF\7-ASUS-GL522VW\SAFE NOTE\Autokey -- 75-GOODSYNC OPTIONS SET_GOODSYNC KEY.txt"
	IfExist, %File_NAME%
	{
		FileGetTime, OutputVar, %File_NAME%, M
		IF FILE_NAME_DATE_MOD_GOODSYNC_KEY_SCRIPT<>%OutputVar%
		{
			FILE_NAME_DATE_MOD_GOODSYNC_KEY_SCRIPT = %OutputVar%
			GOSUB SET_ACTIVATION_BOX_GOODSYNC_TEXT_FILE
			; HERE IT READER FILE GET IN ARRAY
		}
	}
	
		
	WinGet, HWND_ACTIVATION, ID, GoodSync Activation ahk_class #32770
	WinGet, HWND_ACTIVE_WINDOW, ID, A
	MouseGetPos,,,WIN_ID_UNDER_MOUSE_CURSOR
	IF HWND_ACTIVATION<>%WIN_ID_UNDER_MOUSE_CURSOR%
		SET_ACTIVATION__HWND=0
	IF HWND_ACTIVATION<>%HWND_ACTIVE_WINDOW%
		SET_ACTIVATION__HWND=0

	; TOOLTIP % HWND_ACTIVATION "`n" WIN_ID_UNDER_MOUSE_CURSOR "`n" SET_ACTIVATION__HWND
	
	; [ Monday 19:28:20 Pm_09 September 2019 ]
	; LARGER LARGER LOUDER BEER
	
	DetectHiddenWindows, ON
	SetTitleMatchMode 3
	IfWinExist GoodSync Activation ahk_class #32770
	HWND_ACTIVATION=
	WinGet, HWND_ACTIVATION, ID, GoodSync Activation ahk_class #32770
	IF SET_ACTIVATION__HWND<>%HWND_ACTIVATION%
	IF HWND_ACTIVATION>0
	{
		; SET_ACTIVATION__HWND=%HWND_ACTIVATION%  ; ---- DO LATER
		IDX_MAX:=FN_Array_1.MaxIndex()
		IDX_MAX-=3
		IDX_1=-2
		Loop 200
		{
			IF IDX_1>%IDX_MAX%
				BREAK
			IDX_1+=3
			IDX_2=%IDX_1%
			IDX_2+=1
			IDX_3=%IDX_1%
			IDX_3+=2
			
			Element_1 := FN_Array_1[IDX_1]
			Element_2 := FN_Array_1[IDX_2]
			Element_3 := FN_Array_1[IDX_3]

			; TOOLTIP % Element_1 "`n" Element_2 "`n" Element_3 "`n" 
			
			OK_COMPUTER_NAME=
			IF Element_1=*
				OK_COMPUTER_NAME=TRUE
			IF A_ComputerName=%Element_1%
				OK_COMPUTER_NAME=TRUE

			; TOOLTIP % Element_1 "`n" OK_COMPUTER_NAME
			
			; GET AGAIN SPEED REQUIRE MAYBE
			WinGet, HWND_ACTIVATION, ID, GoodSync Activation ahk_class #32770
			WinGet Path, ProcessPath, ahk_id %HWND_ACTIVATION%

			IF OK_COMPUTER_NAME
			IF HWND_ACTIVATION>0
			IF INSTR(Path,Element_2)
			{
				; TOOLTIP % PATH "`n" Element_2
				ControlGettext, OutputVar_2, Edit2, ahk_id %HWND_ACTIVATION%
				IF OutputVar_2<>%Element_3%
				{
					ControlSetText, Edit1,, ahk_id %HWND_ACTIVATION%
					Control, EditPaste, MATTHEW LANCASTER, Edit1, ahk_id %HWND_ACTIVATION%
					ControlSetText, Edit2,, ahk_id %HWND_ACTIVATION%
					Control, EditPaste, %Element_3%, Edit2, ahk_id %HWND_ACTIVATION%
					SoundBeep , 4000 , 100
				}
					
				; THERE TO TYPE OF ACTIVATION FORM
				; ONE IS PULL DOWN MENU ASK
				; ANOTHER COME AS GO TO DO ANYTHING REQUIRE THEM
				; WHEN 2ND ONE COME 
				; IT IS EDIT BOX 4 AND 5
				; I FILL BOTH SO LAZY TO CODE PROPERLY
					
				ControlGettext, OutputVar_2, Edit5, ahk_id %HWND_ACTIVATION%
				IF OutputVar_2<>%Element_3%
				{
					ControlSetText, Edit4,, ahk_id %HWND_ACTIVATION%
					Control, EditPaste, MATTHEW LANCASTER, Edit4, ahk_id %HWND_ACTIVATION%
					ControlSetText, Edit5,, ahk_id %HWND_ACTIVATION%
					Control, EditPaste, %Element_3%, Edit5, ahk_id %HWND_ACTIVATION%
					SoundBeep , 4000 , 100
				}
			}
		}
		
		IF SET_ACTIVATION__HWND<>%HWND_ACTIVATION%
		{
			ControlGettext, OutputVar_2, Edit2, ahk_id %HWND_ACTIVATION%
			IF OutputVar_2=%Element_3%
			{
				SET_ACTIVATION__HWND=%HWND_ACTIVATION%
				RETURN
			}
			ControlGettext, OutputVar_2, Edit5, ahk_id %HWND_ACTIVATION%
			IF OutputVar_2=%Element_3%
			{
				SET_ACTIVATION__HWND=%HWND_ACTIVATION%
				RETURN
			}
		}
	}
}
RETURN

; OPTIONAL SUB ROUTINE
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


TOOLTIP_REMOVER:
	IF TOOLTIP_SET_REMOVE_TIMER
	IF TOOLTIP_SET_REMOVE_TIMER<%A_NOW%
	{
		TOOLTIP
		TOOLTIP_SET_REMOVE_TIMER=
	}
	SETTIMER TOOLTIP_REMOVER,OFF
RETURN


; -------------------------------------------------------------------
TIMER_SUB_GOODSYNC_OPTIONS:

	; IF TOOLTIP_SET_REMOVE_TIMER
	; IF TOOLTIP_SET_REMOVE_TIMER<%A_NOW%
	; {
		; TOOLTIP
		; TOOLTIP_SET_REMOVE_TIMER=
	; }

	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, ON
	SetTitleMatchMode 2  ; Avoids Specify Full path.

	; ---------------------------------------------------------------
	NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D=
	WinGet, HWND_1, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
	IF HWND_1>0
	{
		; -----------------------------------------------------------
		WinGet, HWND_2, ID, A
		HWND_GS_A_PARENT=%HWND_2%
		IF HWND_2<>%HWND_1%
		{
			aParent:=DllCall( "GetParent", UInt, HWND_2) + 0
			HWND_2=%aParent%
			HWND_GS_A_PARENT=%HWND_2%
		}
		; -----------------------------------------------------------
		IF HWND_2=%HWND_1%
		{
			WinGet Path, ProcessPath, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
			IF INSTR(Path,"D:\GoodSync\x64\GoodSync2Go.exe")
				NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D=TRUE
		}
		; -----------------------------------------------------------
	}
	; ---------------------------------------------------------------

	WinGet, HWND_1, ID, ] Options ahk_class #32770
	IF !HWND_1
		RETURN

	WinGet, HWND_1_EXENAME, ProcessName, ahk_id %HWND_1%
	WinGet, HWND_2, ID, A
	IF HWND_1
	IF HWND_2=%HWND_1%
	{
		;---------------------------------------------------------
		; WinGet, OutputVar, ControlList, ahk_id %HWND_1%
		;---------------------------------------------------------
		; Tooltip, % OutputVar ; List All Controls of Active Window
		;---------------------------------------------------------
		; -----------------------------------------------------------
		ControlGet, OutputVar_1, Line, 1, Edit9, ahk_id %HWND_1%
		ControlGettext, OutputVar_2, Button17, ahk_id %HWND_1%
		WinGetTitle OutputVar_3,ahk_id %HWND_1%
		; -----------------------------------------------------------
		SET_GO=FALSE
		; IF A_ComputerName=8-MSI-GP62M-7RD
		IF HWND_1_EXENAME=GoodSync-v10.exe
		{
			IF OutputVar_1=2
				SET_GO=TRUE
			IF OutputVar_1=5
				SET_GO=TRUE
		}
		; -----------------------------------------------------------
		; APPLY TO  -- GoodSync-v10.exe
		; -----------------------------------------------------------
		IF PERIODICALLY_SET_VALUE_HWND<>%HWND_1%
		IF SET_GO=TRUE
		IF HWND_1_EXENAME=GoodSync-v10.exe
		IF OutputVar_1=2
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			ControlSetText, Edit9,, ahk_id %HWND_1%
			TOOLTIP AUTO - SET PERIODIC VALUE
			TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
			TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

			VALUE_PASTE_IN=5
			Control, EditPaste, %VALUE_PASTE_IN%, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
			ControlGet, OutputVar_3, Line, 1, Edit9, ahk_id %HWND_1%
			IF OutputVar_3=%VALUE_PASTE_IN%
			{
				PERIODICALLY_SET_VALUE_HWND=%HWND_1%
				TOOLTIP
			}	
		}
		; -----------------------------------------------------------
		; SET THE Periodically IS NOT SET VALUE WANT
		; -----------------------------------------------------------
		; APPLY TO  -- GoodSync2Go.exe
		; -----------------------------------------------------------
		IF !NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D
		IF PERIODICALLY_SET_VALUE_HWND<>%HWND_1%
		IF SET_GO=TRUE
		IF HWND_1_EXENAME=GoodSync2Go.exe
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			SET_GO=FALSE
			ControlSetText, Edit9,, ahk_id %HWND_1%
			TOOLTIP AUTO - SET PERIODIC VAULE
			TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
			TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS
			
			; REMOVE -- HOW RUDE
			VALUE_PASTE_IN=4
			Control, EditPaste, %VALUE_PASTE_IN%, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100

			ControlGet, OutputVar_3, Line, 1, Edit9, ahk_id %HWND_1%
			IF OutputVar_3=%VALUE_PASTE_IN%
			{
				PERIODICALLY_SET_VALUE_HWND=%HWND_1%
				TOOLTIP
			}	
		}
		; -----------------------------------------------------------
		; APPLY TO  -- D DRIVE GoodSync2Go.exe
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		IF NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D=TRUE
		IF PERIODICALLY_SET_VALUE_HWND<>%HWND_1%
		IF SET_GO=TRUE
		IF HWND_1_EXENAME=GoodSync2Go.exe
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			SET_GO=FALSE
			ControlSetText, Edit9,, ahk_id %HWND_1%
			TOOLTIP AUTO - SET PERIODIC VALUE
			TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
			TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

			VALUE_PASTE_IN=4
			Control, EditPaste, %VALUE_PASTE_IN%, Edit9, ahk_id %HWND_1%
			SoundBeep , 4000 , 100

			ControlGet, OutputVar_3, Line, 1, Edit9, ahk_id %HWND_1%
			IF OutputVar_3=%VALUE_PASTE_IN%
			{
				PERIODICALLY_SET_VALUE_HWND=%HWND_1%
				TOOLTIP
			}	
		}
		
		; -----------------------------------------------------------
		; -----------------------------------------------------------
	

		; -----------------------------------------------------------
		; SET THE PERIODICALLY WHEN COME STRAIGHT IN ON AUTO FORM 
		; WITH A ONE HITTER
		; -----------------------------------------------------------
		IF !NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D
		IF O_HWND_1<>%HWND_1%
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			TOOLTIP AUTO - CHECK PERIODIC TIMER
			TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
			TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

			ControlGet, OutputVar_4, Visible, , Button17, ahk_id %HWND_1%
			ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
			If Status=0
			{
				Control, Check,, Button17, ahk_id %HWND_1%
				IF OutputVar_4=1
				{
					SoundBeep , 4000 , 100
					TOOLTIP
				}
			}
		}
		
		IF NOT_UPDATE_AWFUL_LOT_GOODSYNC2GO_D=TRUE
		IF UNCHECK_PERIODIC_TIMER_HWND<>%HWND_1%
		IF (OutputVar_2="Periodically (On Timer), every")
		{
			TOOLTIP AUTO - UNCHECK PERIODIC TIMER
			TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
			TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

			ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
			If Status=1
			{
				Control, UNCheck,, Button17, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			ControlGet, Status, Checked,, Button17, ahk_id %HWND_1%
			If Status=0
			{
				UNCHECK_PERIODIC_TIMER_HWND=%HWND_1%
				TOOLTIP
			}

		}

		IF CHECK_ESTIMATE_DISK_SPACE_REQUIRED_HWND<>%HWND_1%
		{
			ControlGettext, OutputVar_2, Button35, ahk_id %HWND_1%
			IF (OutputVar_2="Estimate disk space required for Sync")
			{
				TOOLTIP ADAVANCED - CHECK_ESTIMATE_DISK_SPACE_REQUIRED
				TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
				TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

				ControlGet, Status, Checked,, Button35, ahk_id %HWND_1%
				If Status=0
				{
					Control, Check,, Button35, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
				}
				ControlGet, Status, Checked,, Button35, ahk_id %HWND_1%
				If Status=1
				{
					CHECK_ESTIMATE_DISK_SPACE_REQUIRED_HWND=%HWND_1%
					TOOLTIP
				}
			}
		}
		
		IF WAIT_FOR_LOCKS_TO_CLEAR_MINUTE_HWND<>%HWND_1%
		{
			ControlGettext, OutputVar_2, Button23, ahk_id %HWND_1%
			ControlGet, OutputVar_1, Line, 1, Edit12, ahk_id %HWND_1%
			
			If OutputVar_1<>20
			IF OutputVar_2="Wait for Locks to clear, minutes"
			{
				TOOLTIP AUTO - WAIT_FOR_LOCKS_TO_CLEAR_MINUTE_HWND
				TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
				TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

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
			ControlGet, Status, Checked,, Button23, ahk_id %HWND_1%
			If Status=1
			{
				WAIT_FOR_LOCKS_TO_CLEAR_MINUTE_HWND=%HWND_1%
				TOOLTIP
			}
		}
			
		; Text:	Exclude Temporary files and folders
		; -----------------------------------------------------------
		IF TEMPORARY_FILE_HWND<>%HWND_1%
		{
			; UNCHECKER -- IS Status=0
			; ------------------------
			ControlGet, Status, Checked,, Button13, ahk_id %HWND_1%
			If Status=1
			{
				Control, UNCheck,, Button13, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
				TOOLTIP FILTERS - TEMPORARY_FILE
				TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
				TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS
			}
			ControlGet, Status, Checked,, Button13, ahk_id %HWND_1%
			If Status=0
			{
				TEMPORARY_FILE_HWND=%HWND_1%
				TOOLTIP
			}
		}

		
		IF DO_NOT_SYNC_IF_CHANGED_FILES_MORE_THAN_HWND<>%HWND_1%
		{
			ControlGettext, OutputVar_2, Button21, ahk_id %HWND_1%
			ControlGet, OutputVar_1, Line, 1, Edit11, ahk_id %HWND_1%
			
			If OutputVar_1 = 100 ; DEFAULT
			IF OutputVar_2="Do not Sync if changed files more than"
			{
				TOOLTIP AUTO - Do not Sync if changed files more than
				TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
				TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

				ControlSetText, Edit11,, ahk_id %HWND_1%
				VALUE_PASTE_IN=80
				Control, EditPaste, %VALUE_PASTE_IN%, Edit11, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			
			ControlGet, OutputVar_3, Line, 1, Edit11, ahk_id %HWND_1%
			IF OutputVar_3=%VALUE_PASTE_IN%
			{
				DO_NOT_SYNC_IF_CHANGED_FILES_MORE_THAN_HWND=%HWND_1%
				TOOLTIP
			}	
		}
		
		IF HDD_HUBIC_HWND<>%HWND_1%
		{
			WinGetTitle OutputVar_3, ahk_id %HWND_1%
			If INSTR(OutputVar_3,"HDD HUBIC")>0 
			{
				TOOLTIP ---- - HUBIC UNCHECK BUTTON58
				TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
				TOOLTIP_SET_REMOVE_TIMER+=4, SECONDS

				ControlGet, Status, Checked,, Button58, ahk_id %HWND_1%
				If Status=1
				{
					Control, UNCheck,, Button58, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
				}
			}
			ControlGet, Status, Checked,, Button58, ahk_id %HWND_1%
			If Status=0
			{
				HDD_HUBIC_HWND=%HWND_1%
				TOOLTIP
			}
		}

		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; MORE WORK TO DO FROM HERE BEYOND
		; ALL HWND MUST BE STORE EACH FOR CHANGE
		; TO MAKE SMOOTHER NOT PINGER BOX FOR INFO
		; HIGH REPEATER
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		
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
	}                       ; 	IF HWND_2=%HWND_1%


	; ---------------------------------------------------------------
	;Are you sure you want to move _GSDATA_ folder back to the sync folder?
	;This option is for Advanced users and you should read the manual first.
	;Yes
	;No
	; 
	; ---------------------------------------------------------------
	;WinGet, HWND_1, ID, GoodSync Warning ahk_class #32770
	; ---------------------------------------------------------------
	WinGet, HWND_5, ID, GoodSync ahk_class #32770
	IF HWND_5>0 
	{
		WinGet, HWND_7, ID, A
		IF HWND_7=%HWND_5%
		{
			WinGetText OutputVar_3,ahk_id %HWND_5%

			IfInString, OutputVar_3, Are you sure you want to move _GSDATA_
			{
				ControlClick, Button2,ahk_id %HWND_5%
				SoundBeep , 4000 , 100
			}

			IfInString, OutputVar_3, Are you sure you want to keep _GSDATA_
			{
				ControlClick, Button2,ahk_id %HWND_5%
				SoundBeep , 4000 , 100
			}

			WinGet, HWND_5, ID, GoodSync ahk_class #32770
			WinGetText OutputVar_3,ahk_id %HWND_5%

			IfInString, OutputVar_3, Removable drive with volume name
			{
				ControlClick, Yes,ahk_id %HWND_5%
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
				ControlClick, Button3,ahk_id %HWND_5%
				SoundBeep , 4000 , 100
			}
		}
		
		
	}   ;   IF HWND_5>0 

	; IF HWND_1
	; IF HWND_2=%HWND_1%
	; {
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
	; }

	
	O_HWND_1=%HWND_1%
	
	DetectHiddenWindows, % dhw

Return



CHECK_GOODSYNC_NOT_RESPOND:

	; IF !(A_ComputerName = "7-ASUS-GL522VW") 
		; RETURN

	DetectHiddenWindows, OFF

	WinGet, HWND_4, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}

	IF HWND_4>0 
	{
		WinGetTitle, Title_4, ahk_id %HWND_4%
		
		; TOOLTIP % Title_4
		
		IF INSTR(Title_4,"(Not Responding)")=0
			TIMER_NOT_RESPONDING=0
		
		IF INSTR(Title_4,"(Not Responding)")>0
			IF TIMER_NOT_RESPONDING=0
			{
				TIMER_NOT_RESPONDING = % A_Now
				TIMER_NOT_RESPONDING += 20, MINUTES
			}
			IF TIMER_NOT_RESPONDING>0
				IF TIMER_NOT_RESPONDING<%A_Now%
				{
					SoundBeep , 1000 , 100
					SoundBeep , 1500 , 100
					Process, Close, GoodSync-v10.exe
				}	
	}
	IF TIMER_NOT_RESPONDING>0
	{
		TOOLTIP "TIMER_NOT_RESPONDING GOODSYNC" %TIMER_NOT_RESPONDING%
		TOOLTIP_SET_REMOVE_TIMER=%A_NOW%
		TOOLTIP_SET_REMOVE_TIMER+=20, SECONDS
		SETTIMER TOOLTIP_REMOVER,OFF

	}


RETURN



MINIMIZE_AND_RUN_GOODSYNC_V10:

	; GET GOODSYNC HANDLE AND THEN CHECK IF CHANGE DUE TO UPDATE HAPPEN
	; IF HAS AND THEN MINIMIZE ON CERTAIN COMPUTER
	; -------------------------------------------------------------------
	DetectHiddenWindows, ON
	WinGet, HWND_1, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}

	IF HWND_1>0 
		If GOODSYNC_CHECK_CHANGE_OLD_ONE_HANDLE=%HWND_1%
		{
			; ------------------------------------------------------------------------
			; SO SOMETHING TURNED UP I PUT AN _IF_ IN VARIABLE NAME AND WOULDN'T WORK
			; ------------------------------------------------------------------------
			RETURN
		}
	IF HWND_1>0 
		GOODSYNC_CHECK_CHANGE_OLD_ONE_HANDLE=%HWND_1%

	WinGet, HWND_2, ID, ] Options ahk_class #32770
	IF HWND_2>0 
	WinGet, HWND_3, ID, A
	IF HWND_2=%HWND_3%
		RETURN
		
		
	IF HWND_1>0 
	{
		WinGet, PID_1, PID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
		If GOODSYNC_CHECK_CHANGE_OLD_ONE_PID=%PID_1%
			RETURN
	}
	IF HWND_1>0 
		GOODSYNC_CHECK_CHANGE_OLD_ONE_PID=%PID_1%
		
	IF HWND_1>0 
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

			HAS_MIMIMIZE_DO=0
			LOOP, 100
			{
				WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
				If MMX=-1
				{
					HAS_MIMIMIZE_DO=1
				}
				SLEEP 50
			}
		}
	}
		
	DetectHiddenWindows, ON

RETURN


MINIMIZE_AND_RUN_GOODSYNC_2GO:
	
	; GET GOODSYNC HANDLE AND THEN CHECK IF CHANGE DUE TO UPDATE HAPPEN
	; IF HAS AND THEN MINIMIZE ON CERTAIN COMPUTER
	; -------------------------------------------------------------------
	DetectHiddenWindows, ON
	SET_GO_1=0
	IF (A_ComputerName="4-ASUS-GL522VW" and A_UserName="MATT 01")
		SET_GO_1=1
	IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName="MATT 04")
		SET_GO_1=1
		
	IF SET_GO_1=0 
		RETURN

		
	WinGet, HWND_1, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}

	IF HWND_1>0 
		If GOODSYNC_CHECK_CHANGE_OLD_2GO_HANDLE = %HWND_1%
		{
			; ------------------------------------------------------------------------
			; SO SOMETHING TURNED UP I PUT AN _IF_ IN VARIABLE NAME AND WOULDN'T WORK
			; ------------------------------------------------------------------------
			RETURN
		}
	IF HWND_1>0 
		GOODSYNC_CHECK_CHANGE_OLD_2GO_HANDLE=%HWND_1%

	WinGet, HWND_2, ID, ] Options ahk_class #32770
	IF HWND_2>0 
	WinGet, HWND_3, ID, A
	IF HWND_2=%HWND_3%
		RETURN

	IF HWND_1>0 
	{
			WinGet, PID_1, PID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
			If GOODSYNC_CHECK_CHANGE_OLD_2GO_PID=%PID_1%
				RETURN
	}
	IF HWND_1>0 
		GOODSYNC_CHECK_CHANGE_OLD_2GO_PID=%PID_1%
		
	IF HWND_1>0 
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
			SLEEP 200

			HAS_MIMIMIZE_DO=0
			LOOP, 100
			{
				WinGet MMX, MinMax, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
				If MMX=-1
				{
					HAS_MIMIMIZE_DO=1
				}
				SLEEP 50
			}
		}
	}

	DetectHiddenWindows, ON

RETURN


ROUTINE_SPARE_CODIN:
	; ---------------------------------------------------------------
	; PROBLEM ONE COMPUTER MY MSI INTEL 7 WIN 10 
	; HAS THAT PROCESS IS CALLED GoodSync.exe
	; WHILE OTHER COMPUTER USE   GoodSync-v10.exe
	; SO WE __ 19 JAN 2019 13:50
	; ---------------------------------------------------------------
	SET_GO=TRUE
	IFWINNOTEXIST ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
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



;--------------------------------------------------------------------
TIMER_SUB_GOODSYNC_SCRIPT_COMMAND_TO_STOP:
;--------------------------------------------------------------------
	; setTimer TIMER_SUB_GOODSYNC_SCRIPT_COMMAND_TO_STOP, OFF
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, OFF
	SetTitleMatchMode 2  ; NOT specify the full path of the file below.


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





; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------


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

