	 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY.ahk
;# __ 
;# __ Autokey -- 73-MSGBOX COUNTDOWN DELAY.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; CODE FROM SEPARATION ANOTHER MAIN CODE
; MAIN CODER WAS WORK ON A LOT SO MOVE HERE
; -------------------------------------------------------------------
; FROM __ Sat 16-Mar-2019 07:37:08
; TO   __ Sat 16-Mar-2019 08:44:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 02
; -------------------------------------------------------------------
; WAS BIG CODE UPGRADE FEW DAY BEFORE AND TODAY
; ADD ANOTHER THING TO ARRAY AND FINE TUNE
; QUITE HANDY ROUTINE FOR KEEP PROGRAM RUNNING WITH COUNTDOWN POPPER MSGBOX OFF
; WHEN STAND ALONG AND RERUN THE EXE
; SOMETIME SYNC OF EXE BETWEEN NETWORK WHEN USE MY LESS ADVANCE PROGRAM
; CUT THEM UP OR AND COMPILED EXE WITH ERROR SYNC AROUND
; WAIT FOR GOOD ONE TO COME
; LONG WORK OTHER DAY NEW ARRAY SYSTEM FOR VARIABLE WHEN HARD-CODER
; SETTIMER TIMER_VB_EXE_APPLICATION_ERROR_MSGBOX,1000
; HAVE GOT 4 ARRAY SET
; -------------------------------------------------------------------
; FROM __ Mon 26-Aug-2019 09:37:15
; TO   __ Mon 26-Aug-2019 10:41:30
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location On-Line
;--------------------------------------------------------------------
; ----
; Autokey -- 58-Auto Repeat Browser Function Set.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2058-Auto%20Repeat%20Browser%20Function%20Set.ahk
; ----
;# ------------------------------------------------------------------


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
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

IF A_ScriptDir<>C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY
	MSGBOX % "A_ScriptDir IS NOT AT THE DIRECTORY EXPECTOR `n`n" A_ScriptDir "`n" "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY `n`nINCLUDE FILE MIGHT FIND DIFFICULT `nMAKE SURE CHECK YOU ROOT DIRECTORY"

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
#Include %A_ScriptDir%\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, oFF
SetTitleMatchMode 2  ; Partial Path

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6


Secs_MSGBOX_01=18
Secs_MSGBOX_02=5
Secs_MSGBOX_03=50
Secs_MSGBOX_04=5
Secs_MSGBOX_05=
Secs_MSGBOX_06=
Secs_MSGBOX_07=
Secs_MSGBOX_08=0
Secs_MSGBOX_08_RUN_ONCE=FALSE
SHOW_COUNTDOWN_ACTION=FALSE
RELAUNCH_PATH_VAR=

X_COUNT_EXIT=0
MSGBOX_COUNTDOWN_RESTART=""
VAR_WORKER_MSGBOX_DELAY_COUNT=""
OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2

O_DATE_MOD_1=
O_DATE_MOD_2=
O_DATE_MOD_3=

SETTIMER MSGBOX_COUNTDOWN_VB_KEEP_RUNNER_OS_RESTART,1000

SETTIMER MSGBOX_PRESS_FOR_RELOADER,4000
	
SETTIMER TIMER_HOTKEY_VB_CONFIRM_SAVE_AS,1000
	
SETTIMER TIMER_HOTKEY_VB_MSGBOX_MSCOMCTL_OCX,1000	
	
SETTIMER TIMER_MSGBOX_WINDOWS_SCRIPT_HOST_IP_CHANGER,1000

SETTIMER TIMER_MSGBOX_GOODSYNC_EXIT_PROGRAM_ASK_QUESTION_ANOTHR_JOB_RUNNER,1000

SETTIMER TIMER_VB_EXE_APPLICATION_ERROR_MSGBOX,1000

SETTIMER RELOAD_OWN_SCRIPT_IF_INCLUDE_UPDATER,1000
	
SETTIMER SET_OWN_SCRIPT_LESS_PRIORITY_DEPEND_COMPUTER_NAME, 2000

RETURN
; -------------------------------------------------------------------
; END OF INIT DECLARE VARIABLE
; NOW ROUTINE AND FUNCTION
; -------------------------------------------------------------------
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

RELOAD_OWN_SCRIPT_IF_INCLUDE_UPDATER:

	; ---------------------------------------------------------------
	GOSUB SUB_ROUTINE_INCLUDE
	; ---------------------------------------------------------------
	; #Include %A_ScriptDir%\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE_02.ahk
	; ---------------------------------------------------------------
		
RETURN

; -------------------------------------------------------------------
; GOSUB SUB_ROUTINE_INCLUDE
; -------------------------------------------------------------------
#Include %A_ScriptDir%\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE_02.ahk
; -------------------------------------------------------------------


TIMER_HOTKEY_VB_CONFIRM_SAVE_AS:

	; -------------------------------------------
	; [Window Title]
	; Confirm Save As

	; [Content]
	; Shell VBasic 6 Loader.exe already exists.
	; Do you want to replace it?

	; [Yes] [No]
	; -------------------------------------------

	VAR_IN_NAME=Confirm Save As ahk_class #32770
	SetTitleMatchMode 3  ; Specify Full path
	IFWINEXIST %VAR_IN_NAME%
	IFWINEXIST Confirm Save As ahk_exe VB6.EXE
	{
		ControlGetText CONTROL_TEXT,Button1,%VAR_IN_NAME%
		
		STRING_V:=&&Yes  0
		IF INSTR(CONTROL_TEXT,%STRING_V%)>1
		{	
			; NA [v1.0.45+]: May improve reliability. See reliability below.
			ControlClick, Button1,%VAR_IN_NAME%,,,, NA x10 y10 
			SOUNDBEEP 4000,300
			VAR_DONE_ESCAPE_KEY=TRUE
		}
		IF CONTROL_TEXT=&Yes
		{
			Secs_MSGBOX_04=5
			SOUNDBEEP 5000,200
		}

		IF Secs_MSGBOX_04>0 	
			Secs_MSGBOX_04-=1
			
		ControlSetText,Button1,&Yes  %Secs_MSGBOX_04%, %VAR_IN_NAME%
	}
	
	
	DetectHiddenWindows, ON
	SetTitleMatchMode 2
	IFWINEXIST Microsoft Visual Basic [design] ahk_class wndclass_desked_gsk
	IFWINEXIST [design] - [Object Browser] ahk_class wndclass_desked_gsk
	IFWINEXIST Microsoft Visual Basic [design] ahk_exe VB6.EXE
	{
		; ---------------------------------------------------------------------
		; IN VB6 PRESS F2 BY MISTAKE INSTEAD OF JUMP BACK TO WHERE YOU WERE HOTKEY WITH CONTROL KEY
		; AND BRING UP OBJECT BROWSER AND RIGHT CLICK ON IT HIDE YOU HAVE TO DO
		; SOLOTUION HERE
		; ---------------------------------------------------------------------
		ControlGet, OutputVar_4, Visible, , Object Browser, Microsoft Visual Basic ahk_class wndclass_desked_gsk
		IF OutputVar_4=1
			Control, HIDE,, Object Browser, Microsoft Visual Basic ahk_class wndclass_desked_gsk
		SOUNDBEEP 5000,200
	}

RETURN


TIMER_HOTKEY_VB_MSGBOX_MSCOMCTL_OCX:

	; -------------------------------------------
	; Form FindWindow ---
	; Vb6 Loader
	; ---------------------
	; ---------------------------
	; Vb6 Loader
	; ---------------------------
	; #2.2# -- MSCOMCTL.OCX

	; WRONG VERSION -- AUTO CHANGED TO

	; #2.1# -- MSCOMCTL.OCX
	; ---------------------------
	; OK   
	; -------------------------------------------
	;IFWINEXIST Vb6 Loader ahk_exe Shell VBasic 6 Loader.exe
	
	VAR_IN_NAME=Vb6 Loader ahk_class #32770
	SetTitleMatchMode 3  ; Specify Full path
	IFWINEXIST %VAR_IN_NAME%
	{
		ControlGetText CONTROL_TEXT,Button1,%VAR_IN_NAME%
		STRING_V:=OK  0
		IF INSTR(CONTROL_TEXT,%STRING_V%)>1
		{	
			; NA [v1.0.45+]: May improve reliability. See reliability below.
			ControlClick, Button1,%VAR_IN_NAME%,,,, NA x10 y10 
			SOUNDBEEP 4000,300
			VAR_DONE_ESCAPE_KEY=TRUE
		}
		IF CONTROL_TEXT=&Yes
		{
			Secs_MSGBOX_05=4
			SOUNDBEEP 5000,200
		}

		IF Secs_MSGBOX_05>0 	
			Secs_MSGBOX_05-=1
			
		ControlSetText,Button1,OK  %Secs_MSGBOX_05%, %VAR_IN_NAME%
	}

RETURN




TIMER_MSGBOX_WINDOWS_SCRIPT_HOST_IP_CHANGER:

	; -------------------------------------------
	; Form FindWindow ---
	; ---------------------------
	; Windows Script Host
	; ---------------------------
	; C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 23-MY IP.VBS

	; IP Address Has Changed

	; 2019-08-17--03-13-42__7-ASUS-GL522VW__86.169.161.128
	; ---------------------------
	; OK   
	; -------------------------------------------
	;IFWINEXIST Vb6 Loader ahk_exe Shell VBasic 6 Loader.exe
	
	VAR_IN_NAME=Windows Script Host ahk_class #32770
	SetTitleMatchMode 3  ; Specify Full path
	IFWINEXIST %VAR_IN_NAME%
	{
		ControlGettext, MSGBOX_INFO, Static1, %VAR_IN_NAME%
		IF INSTR(MSGBOX_INFO,"VBS 23-MY IP.VBS")
		{
			ControlGetText CONTROL_TEXT,Button1,%VAR_IN_NAME%
			STRING_V:=OK  0
			IF INSTR(CONTROL_TEXT,%STRING_V%)>1
			{	
				; NA [v1.0.45+]: May improve reliability. See reliability below.
				ControlClick, Button1,%VAR_IN_NAME%,,,, NA x10 y10 
				SOUNDBEEP 4000,300
				VAR_DONE_ESCAPE_KEY=TRUE
			}
			IF CONTROL_TEXT=OK
			{
				Secs_MSGBOX_06=40
				SOUNDBEEP 5000,200
			}

			IF Secs_MSGBOX_06>0 	
				Secs_MSGBOX_06-=1
				
			ControlSetText,Button1,OK  %Secs_MSGBOX_06%, %VAR_IN_NAME%
		}
	}

RETURN






MSGBOX_PRESS_FOR_RELOADER:
SET_GO=FALSE
WIN_FIND=Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk ahk_class #32770
IFWINEXIST %WIN_FIND%
	SET_GO=TRUE
WIN_FIND=Autokey -- 32-BRUTE BOOT DOWN.ahk ahk_class #32770
IFWINEXIST %WIN_FIND%
	SET_GO=TRUE
IF SET_GO=TRUE 
{
	ControlGettext, MSGBOX_CONTROL, Static1,%WIN_FIND%
	SET_GO=FALSE
	
	; STRING_SEARCH_VAR=Error:
	; IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		; SET_GO=TRUE
	; STRING_SEARCH_VAR=Warning:
	; IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		; SET_GO=TRUE
	STRING_SEARCH_VAR=Could not close the previous instance of
	IF INSTR(MSGBOX_CONTROL,STRING_SEARCH_VAR)>0
		SET_GO=TRUE

	IF SET_GO=TRUE
	{
		; MSGBOX % MSGBOX_CONTROL
		; WinActivate, %WIN_FIND%
		; SLEEP 200
		ControlClick,&No, %WIN_FIND%
		ControlClick,Button2, %WIN_FIND%
		ControlClick,Button1, %WIN_FIND%
		SOUNDBEEP 500,50
		WinGet, PID_02, PID, %WIN_FIND%
		Process, Close, PID_02

	}
}
RETURN


MSGBOX_COUNTDOWN_VB_KEEP_RUNNER_OS_RESTART:

	

	VAR_WORKER_MSGBOX_DELAY_COUNT_01=VB_KEEP_RUNNER ahk_class #32770
	VAR_WORKER_MSGBOX_DELAY_COUNT_02=ahk_class #32770 ahk_exe WScript.exe
	VAR_WORKER_MSGBOX_DELAY_COUNT_03=Autokey -- ahk_class #32770
	VAR_WORKER_MSGBOX_DELAY_COUNT=

	IF !VAR_WORKER_MSGBOX_DELAY_COUNT
	{
		VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_01%
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		{
			ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			STRING_SEARCH_COUNTDOWN_DELAY=READY FOR OS RESTART
			IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
				VAR_WORKER_MSGBOX_DELAY_COUNT=
		}
		IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}

	IF !VAR_WORKER_MSGBOX_DELAY_COUNT
	{
		VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_02%
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		{
			ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			STRING_SEARCH_COUNTDOWN_DELAY=GoodSync Script Command to Stop
			IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
				VAR_WORKER_MSGBOX_DELAY_COUNT=
		}
		IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}

	IF !VAR_WORKER_MSGBOX_DELAY_COUNT
	{
		VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT_03%
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
		{
			ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			STRING_SEARCH_COUNTDOWN_DELAY=Error at Line
			STRING_SEARCH_COUNTDOWN_DELAY=The program will exit
			IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)=0
				VAR_WORKER_MSGBOX_DELAY_COUNT=
		}
		IFWINNOTEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT%
			VAR_WORKER_MSGBOX_DELAY_COUNT=
	}


	IF !VAR_WORKER_MSGBOX_DELAY_COUNT
		RETURN

	MSGBOX_ACTIVATE=FALSE

	IF OLD_VAR_WORKER_MSGBOX_DELAY_COUNT<>%VAR_WORKER_MSGBOX_DELAY_COUNT%
	{
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
		{
			; VB_KEEP_RUNNER _ REBOOT
			Secs_MSGBOX_01=40
			Secs_MSGBOX_02=5
			MSGBOX_ACTIVATE=TRUE
			X_COUNT_EXIT=0
		}
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
		{
			; GOODSYNC
			Secs_MSGBOX_01=150
			Secs_MSGBOX_02=5
			MSGBOX_ACTIVATE=TRUE
			X_COUNT_EXIT=0
		}
		IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
		{
			; GOODSYNC
			Secs_MSGBOX_01=24
			Secs_MSGBOX_02=5
			MSGBOX_ACTIVATE=TRUE
			X_COUNT_EXIT=0
		}
	}
	OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=%VAR_WORKER_MSGBOX_DELAY_COUNT%

	IF %VAR_WORKER_MSGBOX_DELAY_COUNT%
		IF Secs_MSGBOX_01 > -1
		{
			;IF Mod(Secs_MSGBOX_01, 10)=0 
			;	WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			IF MSGBOX_ACTIVATE=TRUE
			{
				MSGBOX_ACTIVATE=FALSE
				WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			}
			
			ControlGettext, MSGBOX_COUNTDOWN_RESTART, Static1, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			IF INSTR(MSGBOX_COUNTDOWN_RESTART,STRING_SEARCH_COUNTDOWN_DELAY)
			{
				SET_GO=FALSE
				IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
					SET_GO=TRUE
				IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
					SET_GO=TRUE
				IF SET_GO=TRUE
				{
					ControlSetText,Static1,%MSGBOX_COUNTDOWN_RESTART%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
					ControlSetText,Button1,&Yes  %Secs_MSGBOX_01%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
					ControlSetText,Button2,Not, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				}
				SET_GO=FALSE
				IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
					SET_GO=TRUE
				IF SET_GO=TRUE
				{
					; ControlSetText,Static1,%MSGBOX_COUNTDOWN_RESTART%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
					ControlSetText,Button1,&OK  %Secs_MSGBOX_01%, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				}
				
				; #WinActivateForce, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				Secs_MSGBOX_01-=1
				RETURN
			}
			ELSE
			{
				X_COUNT_EXIT=0
				RETURN
			}
		
		
		}

	
	
	Secs_MSGBOX_02-=1
		
	IF Secs_MSGBOX_01=-1
	{
			X_COUNT_EXIT+=1
			IF Mod(X_COUNT_EXIT, 20)=0 
				WinActivate, %VAR_WORKER_MSGBOX_DELAY_COUNT%
			IF INSTR(STRING_SEARCH_COUNTDOWN_DELAY,"GoodSync Script Command to Stop")=0
			{
				SET_GO=FALSE
				IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_01%
					SET_GO=TRUE
				IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_02%
					SET_GO=TRUE
				; ControlGetText CONTROL_TEXT,Button1,, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				;IF INSTR(CONTROL_TEXT,&Yes  0)=0
				;	SET_GO=FALSE

				IF SET_GO=TRUE
				{
					ControlClick, &Yes  0, %VAR_WORKER_MSGBOX_DELAY_COUNT%
					SOUNDBEEP 500,50
					OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2
				}
			}
			IFWINEXIST %VAR_WORKER_MSGBOX_DELAY_COUNT_03%
				; ControlGetText CONTROL_TEXT,Button1,, %VAR_WORKER_MSGBOX_DELAY_COUNT%
				; IF INSTR(CONTROL_TEXT,"OK  0")>0
				{
					ControlClick, OK  0, %VAR_WORKER_MSGBOX_DELAY_COUNT%
					SOUNDBEEP 500,50
					OLD_VAR_WORKER_MSGBOX_DELAY_COUNT=-2
				}
		
	}
Return

; MSGBOX COUNTDOWN DELAY
; -------------------------------------------------------------------
TIMER_MSGBOX_GOODSYNC_EXIT_PROGRAM_ASK_QUESTION_ANOTHR_JOB_RUNNER:

	; -------------------------------------------
	; Form FindWindow ---
	; ---------------------------
	; Windows Script Host
	; ---------------------------
	; C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 23-MY IP.VBS

	; IP Address Has Changed

	; 2019-08-17--03-13-42__7-ASUS-GL522VW__86.169.161.128
	; ---------------------------
	; OK   
	; -------------------------------------------
	;IFWINEXIST Vb6 Loader ahk_exe Shell VBasic 6 Loader.exe
	
	; RETURN

	SetTitleMatchMode 3  ; Specify Full path
	
	SET_GO_GS=FALSE
	VAR_IN_NAME_1=GoodSync ahk_class #32770
	VAR_IN_NAME_2=GoodSync2Go ahk_class #32770
	IFWINEXIST %VAR_IN_NAME_1%
		SET_GO_GS=TRUE
	IFWINEXIST %VAR_IN_NAME_2%
		SET_GO_GS=TRUE
	IFWINEXIST %VAR_IN_NAME_1%
		VAR_IN_NAME=%VAR_IN_NAME_1%
	IFWINEXIST %VAR_IN_NAME_2%
		VAR_IN_NAME=%VAR_IN_NAME_2%

	IFWINEXIST %VAR_IN_NAME%
	{
		ControlGettext, MSGBOX_INFO, Edit1, %VAR_IN_NAME%
		IF INSTR(MSGBOX_INFO,"One or more jobs")
		{
			ControlGetText CONTROL_TEXT,Button2,%VAR_IN_NAME%
			STRING_V:=OK  0
			IF INSTR(CONTROL_TEXT,%STRING_V%)>1
			{	
				; NA [v1.0.45+]: May improve reliability. See reliability below.
				ControlClick, Button2,%VAR_IN_NAME%,,,, NA x10 y10 
				SOUNDBEEP 4000,300
				VAR_DONE_ESCAPE_KEY=TRUE
			}

			
			IF INSTR(CONTROL_TEXT,"OK")>0
			IF StrLen(CONTROL_TEXT)=4
			{
				Secs_MSGBOX_07=20
				SOUNDBEEP 5000,200
			}

			IF Secs_MSGBOX_07>0 	
				Secs_MSGBOX_07-=1
			
			ControlSetText,Button2,OK  %Secs_MSGBOX_07%, %VAR_IN_NAME%
		}
	}

RETURN

; TIMER_VB_EXE_APPLICATION_ERROR_MSGBOX,1000
#Include %A_ScriptDir%\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk




SUB_MESS_SPARE_CODE:
; CoordMode, Mouse, SCREEN
		; #WinActivateForce, VB_KEEP_RUNNER ahk_class #32770
		; WinActivate, VB_KEEP_RUNNER ahk_class #32770
		; WinGetPos, X_2, Y_2, , , VB_KEEP_RUNNER ahk_class #32770
		; ControlGetPos, x, y, w, h, Button1, VB_KEEP_RUNNER ahk_class #32770
		; if Secs_MSGBOX_02>0
			; MouseMove, X+20+X_2, Y+20+Y_2
		
Return





#Include %A_ScriptDir%\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk



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

