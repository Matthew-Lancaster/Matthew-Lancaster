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
; IF ERROR WHILE TRY COPY VB CODE EXE IN GOODSYNC
; THEN DO KILL ALL VBCODE EXE 
; AND RELOAD IT SO GOODSYNC RUN WORKER
; -------------------------------------------------------------------
; NOT REAY YET -- DEBUG TO DO
; -------------------------------------------------------------------

; HERE RUN THE EXE WHEN DATE IS NEWER
; IT NOT RELOADER VIA KILL PROCESS
; IT USED IN GOODSYNC HERE
; VB G7 VB-NT 1X _1_ EXE ONLY Autokey -- 68-GOODSYNC ERROR OVERWRITE SCRIPT
; SCRIPT LINE WITH POST SYNC
; nowait: "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk" JOBNAME %JOBNAME% RESULT %RESULT% ERRORS %ERRORS%
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
; NOT RUNNER AT MO
; WAS IDEA NOT YET COMPLETE
; DEBUG SHOWING WANT TO DO
; -------------------------------------------------------------------
EXITAPP

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, KILL   ALL NETWORK - VB CODE.exe, MenuHandler 
Menu, Tray, Add, RELOAD ALL NETWORK - VB CODE.exe, MenuHandler 
Menu, Tray, Add  ; Creates a separator line.

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

MSGBOX % COMMAND_LINE_JOB_NAME_VALUE

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
	FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe"
	FN_Array_2[ArrayCount]:="VB KEEP RUNNER"
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

	GOSUB TIMER_SUB_EliteSpy

	GOSUB TIMER_SUB_VB_KEEP_RUNNER

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
TIMER_SUB_EliteSpy:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist EliteSpy+ by Andrea
{
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
	IfExist, %FN_VAR%
		{
			Run, %FN_VAR%
		}
}
DetectHiddenWindows, % dhw
Return

; -------------------------------------------------------------------
TIMER_SUB_VB_KEEP_RUNNER:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist VB KEEP RUNNER
{
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe"
	IfExist, %FN_VAR%
		{
			Run, %FN_VAR%
		}
}
DetectHiddenWindows, % dhw
Return

; -------------------------------------------------------------------
TIMER_SUB_CPU_INDIVIDUAL_PROCESS:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
DetectHiddenWindows, ON
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinNotExist INDIVIDUAL PROCESS
{
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe"
	IfExist, %FN_VAR%
		{
			Run, %FN_VAR%
		}
}
DetectHiddenWindows, % dhw
Return



MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	MNU_CODE:=A_ThisMenuItem
	
	; MNU_CODE=RELOAD ALL NETWORK - VB CODE.exe
	; MNU_CODE=KILL   ALL NETWORK - VB CODE.exe
	; TIMER_KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	if MNU_CODE=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	if MNU_CODE=Terminate All AutoHotKey.exe
	{
		Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		
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

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	if MNU_CODE=RELOAD ALL NETWORK - VB CODE.exe
	{
		GOSUB RELOAD_ALL_NET___VB_CODE_EXE_SUB
	}
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	if MNU_CODE=KILL   ALL NETWORK - VB CODE.exe
	{
		GOSUB KILL_ALL_NET_VB_CODE_EXE_01

	}
	
return

RELOAD_ALL_NET___VB_CODE_EXE_SUB:

		FileName_2=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE
		
		ArrayCount = 0
		Loop, Read, C:\NETWORK_COMPUTER_NAME.txt 
		{
			NET_PATH:=A_LoopReadLine
			
			SET_GO=TRUE
			IF INSTR(NET_PATH,"BTHUB")
				SET_GO=FALSE
			IF INSTR(NET_PATH,"NAS-QNAP-ML")
				SET_GO=FALSE
			IF SET_GO=TRUE
			{
				ArrayCount += 1
				Array_NETPATH_01%ArrayCount% = %NET_PATH%
				Array_NETPATH_02%ArrayCount% :=StrReplace(NET_PATH, "-", "_")
				ELEMENT1=\\
				ELEMENT2:=Array_NETPATH_01%ArrayCount%
				ELEMENT3=\
				ELEMENT4:=Array_NETPATH_02%ArrayCount%
				ELEMENT5=%FileName_2%
				NET_PATH:=A_LoopReadLine
				ELEMENT7=_%NET_PATH%.TXT

				Array_FileName%ArrayCount% =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
			}
		}

		Loop %ArrayCount%
		{
			FileDelete, % Array_FileName%A_Index%
			SOUNDBEEP 1000,100
		}
		SOUNDBEEP 2000,100

RETURN
 


KILL_ALL_NET_VB_CODE_EXE_01:
		
		FileName_2=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE
		
		ArrayCount = 0
		Loop, Read, C:\NETWORK_COMPUTER_NAME.txt 
		{
			NET_PATH:=A_LoopReadLine
			
			SET_GO=TRUE
			IF INSTR(NET_PATH,"BTHUB")
				SET_GO=FALSE
			IF INSTR(NET_PATH,"NAS-QNAP-ML")
				SET_GO=FALSE
			IF SET_GO=TRUE
			{
				ArrayCount += 1
				Array_NETPATH_01%ArrayCount% = %NET_PATH%
				Array_NETPATH_02%ArrayCount% :=StrReplace(NET_PATH, "-", "_")
				ELEMENT1=\\
				ELEMENT2:=Array_NETPATH_01%ArrayCount%
				ELEMENT3=\
				ELEMENT4:=Array_NETPATH_02%ArrayCount%
				ELEMENT5=%FileName_2%
				NET_PATH:=A_LoopReadLine
				ELEMENT7=_%NET_PATH%.TXT

				Array_FileName%ArrayCount% =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
			}
		}

		; -----------------------------------------------------------
		; GOT DEBUG ERROR HERE -- CAN'T WORK HERE
		; THE JOB RUN IN TEST BY ITSELF TO RUN SCRIPT
		; AND RETURN THE ERROR IN MSGBOX BELOW
		; [ Wednesday 13:20:40 Pm_10 April 2019 ]
		; -----------------------------------------------------------
		Loop %ArrayCount%
		{
			file := FileOpen(Array_FileName%A_Index%, "w")
			if !IsObject(file)
			{
				MsgBox Can't open`r`n"%FileName%"`r`nfor writing.`r`nAutokey -- 68-GOODSYNC REPORT ERROR OVERWRITE VB.ahk
				return
			}
			TestString := "This is a test string.`r`n"  
			file.Write(TestString)
			file.Close()
			SOUNDBEEP 1000,100
		}
		SOUNDBEEP 2000,100

RETURN


TIMER_KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE:

	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, OFF
	SetTitleMatchMode 2  ; Avoids the need to specify the full path

	FileName_2=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE_%A_ComputerName%.TXT
	
	NET_PATH = %A_ComputerName%

	ELEMENT1=\\
	ELEMENT2=%NET_PATH%
	NET_PATH := StrReplace(NET_PATH, "-", "_")
	ELEMENT3=\
	ELEMENT4=%NET_PATH%
	ELEMENT5=%FileName_2%

	FileName_2 =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%
	
	FileExist_FLAG=FALSE
	if FileExist(FileName_2)
		FileExist_FLAG=TRUE
		
	; IF FileExist_FLAG<>%OLD_FileExist_FLAG%
		IF FileExist(FileName_2)
		{
		
			WINCLOSE EliteSpy+ by Andrea B 2001 __
			WINCLOSE VB KEEP RUNNER
			WINCLOSE INDIVIDUAL PROCESS _ Ver
			
			; Process, Close, VB KEEP RUNNER.exe
		}		

	; IF FileExist_FLAG<>%OLD_FileExist_FLAG%
		IF !FileExist(FileName_2)
		{
		
			GOSUB TIMER_SUB_EliteSpy
			GOSUB TIMER_SUB_VB_KEEP_RUNNER
			; GOSUB TIMER_SUB_CPU_INDIVIDUAL_PROCESS
			
		}		
		
	OLD_FileExist_FLAG=%FileExist_FLAG%

DetectHiddenWindows, % dhw
	
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

