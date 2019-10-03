;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# __ Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Fri 17:20:00 Pm_04 May 2018 ]
;# END         TIME [ Fri 17:40:00 Pm_04 May 2018 ]
;# LAST EDITOR TIME [ Mon 12:15:00 Pm_22 Oct 2018 ]
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Fri 04-May-2018 17:20:00
; TO   TIME __ Fri 04-May-2018 17:40:00 BEGINNER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 002 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Thu 07-Jun-2018 16:21:43
; TO   TIME __ Thu 07-Jun-2018 16:27:00 ADD AHK SCRIPT ALREADY RUNNING
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 003 ---------------------------------------------------------------
; -------------------------------------------------------------------
; NOW WITH MORE INTELLIGENT MODIFIED DATE RELOADER 
; AND USER OF ARRAY'S
; AND RENAME FROM 
; Autokey -- 28-Autohotkeys Set AutoRun.ahk
; TO
; Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
; -------------------------------------------------------------------
; FROM TIME __ Wed 20-Jun-2018 21:41:37
; TO   TIME __ Wed 20-Jun-2018 22:28:00 _ 50 MINUTE LEARN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; 004 ---------------------------------------------------------------
; -------------------------------------------------------------------
; ADD CODE - MORE PRECISE USING EXACT - SetTitleMatchMode 3 ; EXACTING
; DEPENDENT ON VERSION NUMBER OF AHK IS SAME FOREVER UNLESS MODIFY
;
; ADD CODE - BETTER USAGE OF ARRAY - ARRAY HANDLING TOOK ALL THE TIME
; LEARN LOTS MORE LIKE REMOVE AND INSERT UNABLE TO CHANGE AN 
; ARRAY VALUE -- SEEM AN EFFORT AS HAVE TO PUSH WHOLE STACK 
; UP AND DOWN AGAIN _ AT LEAST IT AUTO
; YOU CHANGE CHANGE TH VLAU ETO NOTHING SIMPLE ENOUGH
; BUT ALSO PROBLEM AROSE THAT A SECOND ARRAY I WANTER REQUIRED THAT 
; ALL THE ITEM WERE INSERTED IN WITH NEW CODE DISCOVERY 
; DURING A LOOP LEARN
; " - AutoHotkey v1.1.26.01"
; YOU COULD TRIM VERSION NUMBER DOWN A BIT AND NOT USE EXACTING AGAIN
;
; WELL IT ONLY TOOK 10 MINUTE TO LOCATE VERSION NUMBER AND USE IT ON
;
; -------------------------------------------------------------------
; FROM TIME __ Sat 20-Oct-2018 01:58:30
; TO   TIME __ Sat 20-Oct-2018 04:30:00 _ 2 HOUR 30 MINUTE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 005 ---------------------------------------------------------------
; -------------------------------------------------------------------
; ADD CODE _ MADE A CHANGE _ IF YOU KILL A PROCESS IT WON'T RELOAD 
; BECAUSE THE APP FILE PROGRAM WILL ALREADY BEEN SCANNED IN
; IT ONLY LOOKING FOR DATE CHANGER NOT IF IT DISAPPEARED
; THEN IF WORKING ON SOMETHING APP AHK AND KILL IT WON'T RELOAD AGAIN 
; QUICKLY _ SUSSED THAT LOOP GONE
; -------------------------------------------------------------------
; FROM TIME __ Sun 21-Oct-2018 13:58:58
; TO   TIME __ Sun 21-Oct-2018 14:05:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 007 ---------------------------------------------------------------
; -------------------------------------------------------------------
; ADD CODE _ DON'T RUN UPDATED CODE ALREADY RUNNER WHEN IDLE NOT AT 
; CERTAIN LEVEL
; TIDY SPAGHETTI CODE THAT AREA - UNNECESSARY OVER USE OF VARIABLES
; -------------------------------------------------------------------
; FROM TIME __ Mon 22-Oct-2018 11:07:12
; TO   TIME __ Mon 22-Oct-2018 12:15:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 008 ---------------------------------------------------------------
; -------------------------------------------------------------------
; THE CODE _ HAS WORK IN BRUTE SHUT-DOWN MODE NOW
; READ THE REAM STATEMENT LINE IN CODE
; THERE WAS ERROR WHEN LAUNCHED CODE IS RUN AND SOMETHING TERMINATE THAT 
; THEN ALSO THIS CODE HAS HANDLE ON IT AND TERMINATE ALSO
; MAYBE THE SOLUTION
; MAYBE NOT
; THE IDEA BECAME CLEARER WHEN LOOK AT CODE
; WAS SIMPLE TERMINATE ITSELF BEFORE RUN AGAIN AND THEN NOT ABLE 
; SO A SECONDARY LAUNCHER MAKER _ COMPLETE
; -------------------------------------------------------------------
; FROM TIME __ Sat 02-Mar-2019 21:31:50   
; TO   TIME __ Sat 02-Mar-2019 22:08:00
; TO   TIME __ Sat 02-Mar-2019 23:38:00
; -------------------------------------------------------------------

; CERTAINLY CONCERNED - PRO-CON PRO-CERNED PROCEED

; -------------------------------------------------------------------
; 009 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; NOW INCLUDE -- INCLUDE FILE WHEN UPDATE GO IT UPDATE THE FILE WITH 
; INCLUDER
; NEAR FALL ASLEEP FEW TIME TO GET THIS DONE END OF LONG HARD DAY
; -------------------------------------------------------------------
; Wed 28-Aug-2019 23:07:02
; Thu 29-Aug-2019 00:35:10 -- HOW LONG SMALL CODE WANT BEFORE BEDTIME
; ----------------------------------------------------------------------------


;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk· GitHub
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2028-AUTOHOTKEYS%20SET%20RELOADER.ahk
; ----
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
; #Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------


#InstallKeybdHook  ;A_TimeIdlePhysical ignores mouse clicks/mouse moves


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

SetStoreCapslockMode, off
DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

SoundBeep , 2000 , 20

;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------


GLOBAL SIGNAL_TO_RESTART_HAPPEN
GLOBAL I_COUNT

SIGNAL_TO_RESTART_HAPPEN=FALSE
I_COUNT=0


GLOBAL VAR_A__TimeIdle
GLOBAL OutputVar

VAR_A__TimeIdle=0

; Each array must be initialized before use:
FN_Array_1 := [] 
FN_Array_2 := []
FN_Array_3 := []
FN_Array_4 := []
FN_Array_5 := []
DATE_MOD_Array := []

Element_3 := 
Element_4 := 

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

ArrayCount := 0

ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 75-GOODSYNC OPTIONS SET.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk"	
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer.ahk"
IF OSVER_N_VAR>5
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 02-SAVE AS KEY ENTER.ahk"
}
IF OSVER_N_VAR>5
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk"
}
IF OSVER_N_VAR>5
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk"
}
IF OSVER_N_VAR>=10
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK"
}

ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 50-Check The Capital Lock State.ahk"

; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk"

; ArrayCount += 1
; FN_Array_1[ArrayCount] := 

; WIN XP IS 5
IF OSVER_N_VAR>5
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 02-SAVE AS KEY ENTER.ahk"
}



; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\Program Files (x86)\FileZilla Server\FileZilla Server Interface.exe"
; FN_Array_3[ArrayCount] := "FileZilla Server Main Window"


; ; WIN XP IS 5
; IF OSVER_N_VAR>5
; {
	; ArrayCount += 1
	; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 54-Google Chrome Update Process Killer Stop the Tunisia of Advert.ahk"
; }

; SET_GO=TRUE
; IF (A_ComputerName = "2-ASUS-EEE") 
	; SET_GO=FALSE
; IF (A_ComputerName = "8-MSI-GP62M-7RD")
	; SET_GO=FALSE
; IF SET_GO=TRUE
; {
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 58-AUTO REPEAT BROWSER FUNCTION SET.ahk"
; }



SET_GO=TRUE
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SET_GO=FALSE
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO=FALSE
IF (A_ComputerName = "3-LINDA-PC") 
	SET_GO=FALSE
IF (A_ComputerName = "5-ASUS-P2520LA") 
	SET_GO=FALSE

IF SET_GO=TRUE
{
	ArrayCount += 1
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk"
}



; IF (A_ComputerName = "7-ASUS-GL522VW") 


; -------------------------------------------------------------------
; HERE THE INCLUDE SET PAIR IN THIS EXAMPLE
; WHEN THE INCLUDE DETECT CHANGE DATE FILE
; IT RUN THE ADJACENT ONE LESS APP AHK
; QUICK IMPLEMENT
; -------------------------------------------------------------------
; Wed 28-Aug-2019 23:24:50
; -------------------------------------------------------------------
 
; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY.ahk"
; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk"

; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY_02.ahk"
; ArrayCount += 1
; FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk"




ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER.ahk"



; -------------------------------------------------------------------
; -------------------------------------------------------------------
; END OF ARRAY SETTER
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; ADD THE TERMINATOR VERSION NUMBER AND THEN WE ARE ABLE TO USE EXACT 
; STRING MATCHING IN CASE NOTEPAD HAD IT
; -------------------------------------------------------------------

; WHEN GOT AN ARRAY VALUE 
; YOU CAN SET IT TO NOTHING 
; BUT UNABLE TO TAKE IT VALUE AND PUT BACK AGAIN
; MUST USE REMOVE AND INSERT SAME POSITION AGAIN
;
; MAYBE BABY
; MAYBE BABY

Loop % ArrayCount
{
	AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
	
	; WORKING CODE FOR A REFERENCE IF EVER REMOVE AND INSERT
	; TEMP_VAR_1:=FN_Array_1[A_Index]
	; FN_Array_1.RemoveAt(A_Index)
	; TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
	; TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
	; TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
	; FN_Array_1.InsertAt(A_Index,TEMP_VAR_3)

	TEMP_VAR_1:=FN_Array_1[A_Index]
	TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
	TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
	TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
	; TEMP_VAR_4 -- PATH
	TEMP_VAR_4:=SUBSTR(TEMP_VAR_1,INSTR(TEMP_VAR_1,"\",,0)+1)

	FN_Array_2.InsertAt(A_Index,TEMP_VAR_3)

	FN_Array_4.InsertAt(A_Index,TEMP_VAR_4)
	
	; ---------------------------------------------------------------
	; FINALLY GOT THERE WITH MY ARRAY HANDLE-SHIP
	; NICE TIME
	; [ Saturday 03:54:40 Am_20 October 2018 ]
	; REMOVE AND INSERT IF WANT TO CHANGE
	; LOOK AT HOW BEAUTIFUL OUR CODE ARE
	; ALL THOSE STRUGGLING
	; TOOK SOME TAPPING
	; HOW A LOOK AT MY GOLDEN EYE TIME KEEPER
	; Sat 20-Oct-2018 01:58:30
	; Sat 20-Oct-2018 03:58:00 __ 2 HOUR-ING FOR A TINY BIT OF CODE
	;                             TOOK LONGER TO MAKE IT PROGRAMMER THE LANGUAGE
	;                             BECAUSE I'M A CRIPPLE HURTING COLLAR BONE
	; ---------------------------------------------------------------
	; POINTLESS IN THE END MIGHT AS WELL HOLD TWO ARRAY SET
	; FIND OUT NEW INFO CAN'T PUT VALUE INTO NEW ARRAY AT INDEX
	;
	; LEARN SOMETHING CAN'T DO AS EXPECT HERE CREATE A SECOND ARRAY 
	; BUT UNABLE TO ADD ANYTHING UNLESS INSERT AGAIN
	; ---------------------------------------------------------------
}
  
Loop % ArrayCount
{
	Element := FN_Array_1[A_Index]
	IF !Element
	{
		MSGBOX EMPTY ARRAY VALUE AT HERE CODER`n`nAutokey -- 28-AUTOHOTKEYS SET RELOADER.ahk`n`nQUITER
		Process, Close,% DllCall("GetCurrentProcessId")
	}

}
  
Loop % ArrayCount
{
  Element := FN_Array_1[A_Index]
	IfExist, %Element%
	{
		FileGetTime, OutputVar, %Element%, M
		DATE_MOD_Array[A_Index] := OutputVar
		; MSGBOX % DATE_MOD_Array[A_Index]
	}
}


IF (A_ComputerName = "7-ASUS-GL522VW") 
	SETTIMER TIMER_SUB_FileZilla_Server,10000

IF (A_ComputerName = "7-ASUS-GL522VW") 
	SETTIMER TIMER_SUB_HUBIC_LAUNCHER_DELETER,1000 ; 1 SECOND
	;SETTIMER TIMER_SUB_HUBIC_LAUNCHER_DELETER,600000 ; 10 MINUTER

	
FIRST_RUN=TRUE

SETTIMER TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD,200

SETTIMER TIMER_COULD_NOT_WAIT_MSGBOX_CLOSE,10000

SETTIMER TIMER_CLOSE_DIFFERCULT_TO_SHUTDOWN_PROGRAM_WHEN_RESTART,1000
	
RETURN

; -------------------------------------------------------------------
; END OF INIT CODE
; -------------------------------------------------------------------

TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD:

SET_TIMER=FALSE

Loop % ArrayCount
{

	TT_1_LESS=%A_Index%
	TT_1_LESS-=1
	Element_1 := FN_Array_1[A_Index]
	Element_7 := FN_Array_1[TT_1_LESS]  ; -- LESS 
	Element_2 := DATE_MOD_Array[A_Index]
	Element_3 := FN_Array_2[A_Index]
	Element_5 := FN_Array_2[TT_1_LESS]  ; -- LESS 	
	Element_4 := FN_Array_4[A_Index]
	Element_4 = %Element_4% ahk_class #32770

	RUN_APP_GO=TRUE
	IF WinExist(Element_4)
	{
		; MSGBOX % Element_4
		SETTIMER TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD,40000
		Element_8=%Element_2%
		Element_8+= -1, Days
		DATE_MOD_Array.InsertAt(A_Index,Element_8)
		; MSGBOX % Element_2 " -- " DATE_MOD_Array[A_Index]
		SET_TIMER=TRUE
		RUN_APP_GO=FALSE
	}
	IF SET_TIMER=FALSE
		SETTIMER TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD,4000
	
	IfExist, %Element_1%
		FileGetTime, OutputVar, %Element_1%, M
	
	IF (A_TimeIdlePhysical < 2000 or FIRST_RUN=TRUE)
		RETURN
	
	RUN_APP_HAPPEN_FLAG=FALSE

	IfExist, %Element_1%
		IF INSTR(Element_3,"_INCLUDE.ahk")=0
		IF (!WinExist(Element_3))
			IF RUN_APP_GO=TRUE
			{
				DATE_MOD_Array[A_Index] := OutputVar
				GOSUB RUN_THE_APP
				RUN_APP_HAPPEN_FLAG=TRUE
			}
	; ---------------------------------------------------------------
	; NOW INCLUDE -- INCLUDE FILE WHEN UPDATE GO IT UPDATE THE FILE WITH 
	; INCLUDER
	; NEAR FALL ASLEEP FEW TIME TO GET THIS DONE END OF LONG HARD DAY
	; Thu 29-Aug-2019 00:35:10
	; ---------------------------------------------------------------
	IfExist, %Element_1%
		IF INSTR(Element_3,"_INCLUDE.ahk")>0
		IF (!WinExist(Element_5))
			IF RUN_APP_GO=TRUE
			{
				DATE_MOD_Array[A_Index] := OutputVar
				GOSUB RUN_THE_APP
				RUN_APP_HAPPEN_FLAG=TRUE
			}


			
	; IF RUN_APP_HAPPEN_FLAG=FALSE
		; IF OutputVar<>%Element_2%
			; IF RUN_APP_GO=TRUE
			; {
				; ; -----------------------------------------------------------
				; ; PUT AN IDLE DELAY HERE CAN'T HAVE AHK APP THAT ARE STOP RUN
				; ; IMMEDIATELY AGAIN
				; ; -----------------------------------------------------------
				; ; IF (A_TimeIdle > 1000)
				; IF (A_TimeIdlePhysical > 2000 or FIRST_RUN=TRUE)
				; {
					; DATE_MOD_Array[A_Index] := OutputVar
					; GOSUB RUN_THE_APP
				; }
			; }
}

FIRST_RUN=FALSE

RETURN

RUN_THE_APP:

	
	; ----------------------------------------------------------
	; NEW CODE FOR INCLUDE FILE -- DON'T RUN THE INCLUDE BUT 
	; RUN THE ALTERNATIVE PROGRAM THAT INCLUDE IS WITH OR MANY
	; ----------------------------------------------------------
	if INSTR(Element_3,"_INCLUDE.ahk")
	{
	
		; MSGBOX %Element_5%
		WinGet, PID_1, PID, %Element_5% ahk_class AutoHotkey
		IF PID_1>0 
		{
			Process, Close,% PID_1
		}
		Run, %Element_5%
		SOUNDBEEP, 1500,100
		RETURN	
	}

	
	; ----------------------------------------------------------
	; FOUND ANSWER LOOK AT CODE HERE
	; REQUIRE SOME SORT OF HOT SWAP LAUNCHER FOR ITSELF OWN CODE
	; ----------------------------------------------------------
	if INSTR(Element_3,"Autokey -- 28-AUTOHOTKEYS SET RELOADER")
	{
		Run, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
		Process, Close,% DllCall("GetCurrentProcessId")
		SOUNDBEEP, 1500,100
		RETURN
	}
	
	; FIND WINDOW -- Element_3
	; ------------------------
	IF INSTR(Element_3,"Autokey -- 32-BRUTE BOOT DOWN")
	{
		LOOP
		{
			WinGet, PID_1, PID, %Element_3% ahk_class AutoHotkey
			IF PID_1>0 
			{
				Process, Close,% PID_1
				; MSGBOX %Element_3% ahk_class AutoHotkey 
				; MSGBOX % PID_1
			}
			IF !PID_1
				BREAK
		}
		; FIND WINDOW -- Element_3
		; ------------------------
		
		Run, %Element_1%
		SOUNDBEEP, 1500,100
		RETURN
	}

	; Loop % 500
	; {
		; PID_1=
		; PID_2=
		; WinGet, PID_1, PID, %Element_4% ahk_class #32770
		; IF PID_1>0 
		; {
			; ; MSGBOX %PID_1% " -- " %Element_4%
			; ; Process, Close,% PID_1
			; PID_1=
		; }
		; WinGet, PID_2, PID, %Element_3% ahk_class AutoHotkey
		; IF PID_1>0 
		; {
			; Process, Close,% PID_2
			; ; MSGBOX % PID_1 " -- " PID_2
		; }
		; if (!PID_1 and !PID_2)
			; BREAK
	; }
		
	SoundBeep , 2000 , 20
	Run, %Element_1%

	
	; TOOLTIP % Element_1

RETURN

; WinGetTitle, Titel, ahk_id %ID%

; SkriptPath := RegExReplace(Titel, " - AutoHotkey v" A_AhkVersion )
; SplitPath,SkriptPath,KurzName

; WinGet, PID, PID, %SkriptPath% ahk_class AutoHotkey





TIMER_CLOSE_DIFFERCULT_TO_SHUTDOWN_PROGRAM_WHEN_RESTART:
	IfWinExist End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
	{
		IF (OSVER_N_VAR = 10) ; WIN 10
		{
			fVisible=0
			AppVisibility := ComObjCreate(CLSID_AppVisibility := "{7E5FE3D9-985F-4908-91F9-EE19F9FD1514}", IID_IAppVisibility := "{2246EA2D-CAEA-4444-A3C4-6DE827E44313}")
			if (DllCall(NumGet(NumGet(AppVisibility+0)+4*A_PtrSize), "Ptr", AppVisibility, "Int*", fVisible) >= 0)
			IF fVisible=1 
			{
				Send {LWin}
			}
		}

		Sleep 4000
		SoundBeep , 2500 , 100
		WINACTIVATE, End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
		ControlClick, &End Now, End Program - CSR_SYNCML_CLASS_1EF5ED00AB77
	}	
RETURN
	
;SETTIMER TIMER_COULD_NOT_WAIT_MSGBOX_CLOSE,10000

TIMER_COULD_NOT_WAIT_MSGBOX_CLOSE:
	; ---------------------------------------------------------------
	; THIS ROUTINE CODE USER IN TWO PROJECT 
	; MIGHT AS WELL TWO AS ONE CLOSED DOWN WHILE WORK CODE
	; OTHER KEEP STUFF RUNNING
	; ---------------------------------------------------------------
	; 01 Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
	; 02 Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
	; ---------------------------------------------------------------

	LINE_CHECKER_1=Could not close the previous instance of this script.
	LINE_CHECKER_2=Keep waiting?

	DetectHiddenWindows, ON
	SetTitleMatchMode 2  ; Specify PARTIAL path
	VAR_GET:=WINEXIST("Autokey ahk_class #32770")
	IF !VAR_GET
		RETURN 
	LOOP 2 {
		IF A_Index=1 
			STATIC_CONTROL=Static1
		IF A_Index=2
			STATIC_CONTROL=Static2
		WinGet, list, List, Autokey ahk_class #32770
		Loop %list% {
			hwnd := list%A_Index%
			ControlGettext, OutVar_2, %STATIC_CONTROL%, ahk_id %hwnd%
			; TOOLTIP % " -- " OutVar_2

			ControlGettext, OutVar_2, %STATIC_CONTROL%, ahk_id %hwnd%
			IF INSTR(OutVar_2,LINE_CHECKER_1)>0
			IF INSTR(OutVar_2,LINE_CHECKER_2)>0
			{
				SOUNDBEEP 4000,100
				ControlClick, Button2, ahk_id %hwnd%,,,, NA x10 y10
			}
			; -------------------------------------------------------
			; GET THE CONTROL AGAIN BECAUSE IF SUCCESSFULLY 
			; KILL 1ST TIME DON'T WANT TO HITT ON NOTHING
			; BETTER QUICKER METHOD CHECK -- WANT
			; IF CHECK EASIER REMOVE TWO LINE SET OF CODE TO ONE
			; -------------------------------------------------------
			ControlGettext, OutVar_2, %STATIC_CONTROL%, ahk_id %hwnd%
			IF INSTR(OutVar_2,LINE_CHECKER_1)>0
			IF INSTR(OutVar_2,LINE_CHECKER_2)>0
			{
				SOUNDBEEP 4000,100
				ControlClick, Button2, ahk_id %hwnd%
			}
		}
	}
RETURN



TIMER_SUB_HUBIC_LAUNCHER_DELETER:

	SETTIMER TIMER_SUB_HUBIC_LAUNCHER_DELETER,600000 ; 10 MINUTER

	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 34-HUBIC DELETE-ER.VBS"
	
	; Run, "%FN_VAR%" , , HIDE
	Run, "%FN_VAR%" /QUITE_MODE, , HIDE

RETURN




TIMER_SUB_FileZilla_Server:



IfWinNotExist, AHK_CLASS FileZilla Server Main Window
{
	Process, Exist, FileZilla Server.exe
	If Not ErrorLevel
	{
		SoundBeep , 2500 , 20
		RunWait,sc start "FileZilla Server" ; CHECK THE SERVICE IS RUNNING AND RUN IT

	}
	Process, Exist, FileZilla Server Interface.exe
	If Not ErrorLevel
	{
		FN_VAR:="C:\Program Files (x86)\FileZilla Server\FileZilla Server Interface.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 20
			Run, "%FN_VAR%" , , HIDE
		}
	}
}
RETURN



MenuHandler:
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02-INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03-INCLUDE MENU 03 of 03.ahk



;# ------------------------------------------------------------------
; USUAL END BLOCK OF CODE TO HELP EXIT ROUTINE
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

    if ExitReason in Logoff,Shutdown
    {
		
		SIGNAL_TO_RESTART_HAPPEN=TRUE
	
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
		
		
		SoundBeep , 2000 , 20
		SENDINPUT {ESC}
		
		SoundBeep , 2500 , 20
		; KILL ITSELF
		Process, Close,% DllCall("GetCurrentProcessId")

		
		;---------------------------------------------------------------------
		; THE RETURN 1 IS STILL USED BUT THE APP CLOSES ITSELF 
		; PROBABLY BECAUSE THE COUPLE OF PROGRAM ARE FOUND NOT TO EXIST
		; BUT EVEN WITH A MSGBOX IN THE WAY IT STILL GET CLOSE
		;---------------------------------------------------------------------
		
		; IF EXIT_APP_VAR=FALSE 
			; {
			; ; MSGBOX % EXIT_APP_VAR
			; return 1
			; }
	}
	
	
	; ---------------------------------------------------------------
	; PROCESS SCRIPT HERE WILL ONLY CLOSE BY FORCE KILL _ DONE WITHIN
	; ---------------------------------------------------------------
	; WE USE THIS CODE HERE
	; Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
	; BECAUSE OF THE HIGH RATE OF THIS CODE CLOSING DOWN 
	; ESPECIALLY ON WINDOWS-XP COMPUTER 
	; AND MOST LIKELY REASON BECAUSE 
	; THE RE-LAUNCHER MAYBE TERMINATING A SCRIPT PREVIOUSLY LOADED BY IT OWN
	; SOME PROGRAM THE LAUNCH A SCRIPT DON'T LIKE IT WITH ANTHER SCRIPT IS PROCESS KILLED OFF
	; AND THEN IT END ITSELF
	; SO OUR CONQUERED THAT BUT ALLOW IT ONLY BRUTE SHUT-DOWN OF PROCESS
	; NONE WAY TO EXIT ANY-MORE
	; IT HELP WHEN WRITE CODE AND NETWORK IT OVER TO ANOTHER COMPUTER
	; LIKE THE THEME IS TODAY
	; WITH THE CODE STOP TOO MANY REPEAT LOT OF CODE WHILE AT FAULT NOT START UP PROPER
	; OR FAULT WHILE GO
	; THE TRAY MENU ITEM ARE THERE NOW TO EXIT
	; ---------------------------------------------------------------

	RETURN 1


	
	; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
		; THIS ROUTINE WON'T GET CALLED UNLESS ExitFunc HAS RETURN CLEAR TO CLOSE PREVENT WITH RETURN 1
		SoundBeep , 2500 , 20
		SoundBeep , 3000 , 20
		
		; MsgBox, MyObject is cleaning up prior to exiting...

		;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}

; -----------------------------------------------------------------
; exit the app
