;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk
;# __ 
;# __ Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ Mon 04-Mar-2019 05:19:59
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
; HERE RUN THE EXE WHEN DATE IS NEWER
; IT NOT RELOADER VIA KILL PROCESS
; IT USED IN GOODSYNC HERE
; VB G7 VB-NT 1X _1_ EXE ONLY Autokey -- 68-GOODSYNC ERROR OVERWRITE SCRIPT
; SCRIPT LINE WITH POST SYNC
; nowait: "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk" JOBNAME %JOBNAME% RESULT %RESULT% ERRORS %ERRORS%
; -------------------------------------------------------------------
; FROM __ Mon 04-Mar-2019 05:19:59
; TO   __ Mon 04-Mar-2019 05:19:59 __ 
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk



; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, OFF
SetTitleMatchMode 2  ; Specify Full path

GLOBAL OSVER_N_VAR

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
; OSVER_N_VAR:=a_osversion
; IF INSTR(a_osversion,".")>0
	; OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
; IF OSVER_N_VAR=WIN_XP
	; OSVER_N_VAR=5
; IF OSVER_N_VAR=WIN_7
	; OSVER_N_VAR=6

; SET_GO=TRUE
; IF A_ComputerName=1-ASUS-X5DIJ
	; SET_GO=FALSE
; IF A_ComputerName=2-ASUS-EEE
	; SET_GO=FALSE


; Each array must be initialized before use:
FN_Array_1 := []
FN_Array_2 := []
FN_Array_3 := []
DATE_MOD_Array := []

ArrayCount := 0
ArrayCount += 1

X=0

; SETTIMER TIMER_KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE,2000

; -------------------------------------------------------------------
; SCRIPT LINE USE IN GOODSYNC 3 OF 1 OF 3 -- LAST OF SCRIPT POST-SYNC
; -------------------------------------------------------------------
; "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk" JOBNAME %JOBNAME% RESULT %RESLUT% ERRORS %ERRORS%
; -------------------------------------------------------------------

COMMAND_LINE:= % DllCall( "GetCommandLine", "str" )

; -------------------------------------------------------------------
; EXAMPLE #1 OF GIVEN COMMAND LINE
; -------------------------------------------------------------------
; COMMAND_LINE="C:\Program Files\AutoHotkey\AutoHotkey.exe" "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk" JOBNAME "VB G7 VB-NT 1X _1_ QUICKER _NEW_GO GSD (1)" RESULT "" ERRORS 1
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; EXAMPLE #2 OF GIVEN COMMAND LINE
; -------------------------------------------------------------------
; "C:\Program Files\AutoHotkey\AutoHotkey.exe" "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk" JOBNAME "VB G7 VB-NT 1X _1_ QUICKER _NEW_GO GSD (1)" RESULT "Please click Analyze again because another job has Synced and it invalidated analysis results of this job.
; State file: //7-ASUS-GL522VW/7_ASUS_GL522VW_02_D_DRIVE/VB6/VB-NT/00_Best_VB_01/Clipboard Logger/_gsdata_/_file_state_v4._gs
; New: [size=197,794 time=2019-03-04 05:55:35] != 
; Old: [size=197,794 time=2019-03-04 05:55:05]
; " ERRORS 0

; -------------------------------------------------------------------
; EXAMPLE #* LOOK OUT FOR MORE EXAMPLE OF GIVEN COMMAND LINE 
; TYPE OF RESULT FIND
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; GET INTO THE CODE ABOUT TO BEGIN
; -------------------------------------------------------------------

COMMAND_LINE_JOB_NAME:=SubStr(COMMAND_LINE, INSTR(COMMAND_LINE," JOBNAME "))
COMMAND_LINE_JOB_NAME_VALUE:=SubStr(COMMAND_LINE_JOB_NAME,11)
COMMAND_LINE_JOB_NAME_VALUE:=SubStr(COMMAND_LINE_JOB_NAME_VALUE,1,INSTR(COMMAND_LINE_JOB_NAME_VALUE," RESULT")-2)

; MSGBOX % COMMAND_LINE_JOB_NAME_VALUE

GoodSync_GSync:="C:\Program Files\Siber Systems\GoodSync\gsync.exe"

GoodSync_GSync_BATCH:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 70-GOODSYNC REPORT ERROR RUN GOODSYNC WITH PARAM.BAT"


COMMAND_LINE_ERROR:=SubStr(COMMAND_LINE, INSTR(COMMAND_LINE,"ERRORS"))
COMMAND_LINE_ERROR_VALUE:=SubStr(COMMAND_LINE_ERROR,8)

IF COMMAND_LINE_ERROR_VALUE>0 
{
	SETTIMER TIMER_KILL_ALL_NETWORK_VB_CODE_EXE_02, 4000
}
ELSE
{
	EXITAPP
}

SETTIMER TIMER_TO_QUIT_KILL_OWN, 2400000    ; 40 min = 2400000 ms
	
RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------

TIMER_TO_QUIT_KILL_OWN:

	; SET THE VB CODE FLAG TO RELOAD ALL THE VB APP BEFORE QUIT HER
	; -------------------------------------------------------------
	GOSUB TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD

	EXITAPP

RETURN

TIMER_KILL_ALL_NETWORK_VB_CODE_EXE_02:

	GOSUB KILL_ALL_NET_VB_CODE_EXE_01

	SETTIMER TIMER_KILL_ALL_NETWORK_VB_CODE_EXE_02,OFF

	SETTIMER TIMER_CHECK_DATE_VB_PROJECT_FOR_UPDATED,2000

RETURN

TIMER_CHECK_DATE_VB_PROJECT_FOR_UPDATED:

	; ---------------------------------------------------------------
	; IF AN UPDATE OF VB CODE EXE HAPPEN -- MODIFIED DATE
	; IT MEAN RESULT GOOD AND CODE HERE IS PACK UP QUICK 
	; RATHER THAN TIMER WAIT BUT NOTHING RESULT HAPPEN
	; THE RETRY OF RUN GOODSYNC AT THE JOB NAME WOULD HAPPEN
	; SHORTLY AFTER PROGRAM LOAD TO KICK START ALONG QUICKER
	; THAT BE WHEN ALL APP BEEN SHUT DOWN TO UN-BLOCKER 
	; THAT HOW FINDING RESULT
	; HA HA
	; ---------------------------------------------------------------

	; FILL UP THE ARRAY HOW EVER MANY YOU WANT TO
	; ---------------------------------------------------------------
	ArrayCount=0
	
	ArrayCount+=1
	FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
	FN_Array_2[ArrayCount]:="EliteSpy+ by Andrea"
	ArrayCount+=1
	FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
	FN_Array_2[ArrayCount]:="VB_KEEP_RUNNER"
	ArrayCount+=1
	FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe"
	FN_Array_2[ArrayCount]:="INDIVIDUAL PROCESS"

	ArrayCount+=1
	FN_Array_1[ArrayCount]:=""
	FN_Array_2[ArrayCount]:=""
	
	; ---------------------------------------------------------------
	; GOT SOME EMPTY LINE ABOVE FOR QUICK GIVE OVER
	; SO REDUCE COUNTER FOR TOP OF ARRAY BY EMPTY LINE DETECTION
	; SAVER PUT REM LINE
	; ---------------------------------------------------------------
	Loop % ArrayCount
	{
		Element := FN_Array_1[A_Index]
		IF !Element
		{
			ArrayCount-=1
		}
	}
	
	Loop % ArrayCount
	{
		Element := FN_Array_1[A_Index]
		IfExist, %Element%
		{
			FileGetTime, OutputVar, %Element%, M
			DATE_MOD_Array[A_Index] := OutputVar
		}
	}

	SETTIMER TIMER_CHECK_DATE_VB_PROJECT_FOR_UPDATED,OFF
	SETTIMER TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD,4000

RETURN

TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD:

; -------------------------------------------------------------------
; TEST ESCAPEE
; -------------------------------------------------------------------
; X+=1
; IF X>5
; {
	; DATE_MOD_Array[2]:=0
; }
; -------------------------------------------------------------------

Loop % ArrayCount
{
	Element_1 := FN_Array_1[A_Index]
	Element_2 := DATE_MOD_Array[A_Index]
	Element_3 := FN_Array_2[A_Index]

	IfExist, %Element_1%
		FileGetTime, OutputVar, %Element_1%, M
	
	IF OutputVar<>%Element_2%
	{
		SETTIMER TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD, OFF ; HERE ROUTINE
		SETTIMER TIMER_SUB_RELOAD_PROCESS_ARRAY,4000     ;
	}
}
RETURN


TIMER_SUB_RELOAD_PROCESS_ARRAY:
	
	SETTIMER TIMER_SUB_RELOAD_PROCESS_ARRAY,OFF  ; -- HERE ROUTINE

	GOSUB RELOAD_ALL_NET___VB_CODE_EXE_SUB

	GOSUB TIMER_SUB_EliteSpy_2

	GOSUB TIMER_SUB_VB_KEEP_RUNNER_2

	; -------------------------------------------------------------------
	; GOOD WORKER DON'T STAY AT FINISHER
	; -------------------------------------------------------------------
	; RUN, "%GoodSync_GSync%" sync "%COMMAND_LINE_JOB_NAME_VALUE%"
	
	; -------------------------------------------------------------------
	; GOOD WORKER DOES STAY AT FINISHER
	; -------------------------------------------------------------------
	SOUNDBEEP 1500,100
	RUN, "%GoodSync_GSync_BATCH%" sync "%COMMAND_LINE_JOB_NAME_VALUE%"
	
	EXITAPP
	
RETURN



; -------------------------------------------------------------------
TIMER_SUB_EliteSpy_2:
; -------------------------------------------------------------------
DHW_2 := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist EliteSpy+ by Andrea
{
	SoundBeep , 3000 , 100
	FN_VAR_2:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
	IfExist, %FN_VAR_2%
		{
			Run, %FN_VAR_2%
		}
}
DetectHiddenWindows, % DHW_2
Return

; -------------------------------------------------------------------
TIMER_SUB_VB_KEEP_RUNNER_2:
; -------------------------------------------------------------------
DHW := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist VB_KEEP_RUNNER
{
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
	IfExist, %FN_VAR%
		{
			Run, %FN_VAR%
		}
}
DetectHiddenWindows, % DHW
Return





#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


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

