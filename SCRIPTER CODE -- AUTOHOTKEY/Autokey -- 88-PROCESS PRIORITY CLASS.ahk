;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# __ Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Sun 28-Apr-2019 14:38:04 ]
;# LAST EDIT TIME [ Tue 20-Aug-2019 15:40:00 ]
;# __ 
;  =============================================================

; ------------------------------------------------------------------
; Location Internet
;---------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL_2.ahk
; ----
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; -------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
#NoTrayIcon

; -------------------------------------------------------------------
; #Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
; OnExit("ExitFunc")

; Register an object to be called on exit:
; OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; ---------------------------------------------------------------
; ---------------------------------------------------------------

; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk



DetectHiddenWindows, on
; SetStoreCapslockMode, off

; SoundBeep , 2000 , 100
; SoundBeep , 2500 , 100



; OSVER_N_VAR:=a_osversion
; IF INSTR(a_osversion,".")>0
	; OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
; IF OSVER_N_VAR=WIN_XP
	; OSVER_N_VAR=5
; IF OSVER_N_VAR=WIN_7
	; OSVER_N_VAR=6
; IF OSVER_N_VAR=10
; 	WANT_GO=TRUE

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; WANT_GO=
; IF (A_ComputerName<>"7-ASUS-GL522VW")
	; WANT_GO=TRUE
; IF (A_ComputerName<>"4-ASUS-GL522VW")
	; WANT_GO=TRUE


GOSUB PROCESS_PRIORITY_CLASS_MODIFY

	
RETURN


; -------------------------------------------------------------------
; CODE IS RUN TIMER FROM HERE
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
PROCESS_PRIORITY_CLASS_MODIFY:

	; ----
	; Command-line arguments - Rosetta Code
	; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
	; ----
	; AutoHotkey
	; From the AutoHotkey documentation: "The script sees incoming parameters as the variables %1%, %2%, 
	; and so on. In addition, %0% contains the number of parameters passed (0 if none). "
	Command_Params=
	
	PROCESS_CLASS=
	PROCESS_PID=
	PARAM_SET=
	
	FOR n, param in A_Args  ; For each parameter:
	{
		PARAM_SET=%PARAM_SET%%A_Space%%PARAM%
	}	

	; PARAM_SET=Explorer.exe /e,D:\DSC\2015+SONY\2020 CyberShot HX60V\MP4

	CABINETWCLASS_SET_1=
	CABINETWCLASS_SET_2=
	WinGet, List, List, ahk_class CabinetWClass
	Loop %List%  
	{ 
		GL:=List%A_Index% 
		CABINETWCLASS_SET_1=%CABINETWCLASS_SET_1%%A_Space%%GL% 
	}

	RUN, %PARAM_SET%,,,PROCESS_PID
	Process, Priority, %PROCESS_PID%, High

	FIND_DONE=
	LOOP 
	{
		WinGet, List, List, ahk_class CabinetWClass
		Loop %List%  
		{ 
			GL:=List%A_Index% 
			CABINETWCLASS_SET_2=%CABINETWCLASS_SET_2%%A_Space%%GL%
			IF INSTR(CABINETWCLASS_SET_1,GL)=0
				FIND_DONE:=List%A_Index%
			IF FIND_DONE
				BREAK
		}
			IF FIND_DONE
				BREAK
	}
	
	IF FIND_DONE
	{
		WinGet, PROCESS_PID, PID, AHK_ID %FIND_DONE%
		Process, Priority, %PROCESS_PID%, High
	}
	
	PAUSE
	
RETURN


