;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 89-WIN EXIST RETURN EXIT CODE.ahk
;# __ 
;# __ Autokey -- 89-WIN EXIST RETURN EXIT CODE.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME    [ Sat 08-Feb-2020 12:27:51 ]
;# LAST EDIT TIME #1 [ Sat 08-Feb-2020 13:58:00 ]
;# LAST EDIT TIME #2 [ Sat 08-Feb-2020 14:25:00 ]
;# __ 
;  =============================================================

; ------------------------------------------------------------------
; Location Internet
;---------------------------------------------------------------------
; Matthew-Lancaster/Autokey -- 89-WIN EXIST RETURN EXIT CODE.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL_2.ahk
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; ----------------------------------------------------------------
; -- 00* SESSION SIXTH
; WORK TODAY
; ----------------------------------------------------------------
; ADD ROUTINE TO CHECK IF THIS PROJECT IS ALREADY RUN 
; BY ISSUE A AUTOHOTKEYS FIND WINDOW ROUTINE THAT PASS BY PARAMETER
; TAKE A SPECIAL TRICK PASS THE PARAMETER LINE WITH * FILL WHERE SPACE
; EXTRACT OTHER END -- LIKE CONCENTRATE ORANGE -- TETRA
; USELESS OTHERWISE
; ----------------------------------------------------------------
; A TON OF CODER AND LITTLE ROUTINE SIZE AFTER COMPACTOR 
; OF LEFT AFTER 
; NOT FIRST TIME DISCOVER THE OPERATION 
; FIRST TIME MAYBE FROM DOS TO AUTOHOTKEYS AND BACK AGAIN 
; EXIT CODE ERRORLEVEL AND PASS PARAMETER 
; ALWAYS PASS PARAMETER WITH * NOT ANY OTHER CHARACTER LIKE DASH -
; OR YOUR HEADACHE HURTING
; ----------------------------------------------------------------
; CODE THAT RUN
; PARAM_SET=BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT
; ----------------------------------------------------------------
; -- 
; Sat 08-Feb-2020 12:00:22
; Sat 08-Feb-2020 14:25:00 -- 2 HOUR 25 MINUTE
; ----------------------------------------------------------------

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

; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk

DetectHiddenWindows, on


GOSUB FIND_WINDOW_FROM_PARAM_AND_EXIT_AFTER_ERROR_CODE

RETURN


FIND_WINDOW_FROM_PARAM_AND_EXIT_AFTER_ERROR_CODE:

	; ----
	; Command-line arguments - Rosetta Code
	; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
	; ----
	; AutoHotkey
	; From the AutoHotkey documentation: "The script sees incoming parameters as the variables %1%, %2%, 
	; and so on. In addition, %0% contains the number of parameters passed (0 if none). "
	Command_Params=
	PARAM_SET=
	
	FOR n, param in A_Args  ; For each parameter:
	{
		PARAM_SET=%PARAM_SET%%A_Space%%PARAM%
	}	

	PARAM_SET:=StrReplace(PARAM_SET, "*", " ")
		
	; PARAM_SET=Explorer.exe /e,D:\DSC\2015+SONY\2020 CyberShot HX60V\MP4
	; PARAM_SET=BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT
	
	CABINETWCLASS_SET_1=
	FIND_TT=0
	
	WinGet, List, List, ahk_class ConsoleWindowClass
	Loop %List%  
	{ 
		HWND_1:=List%A_Index%
		WinGetTitle, WName, ahk_id %HWND_1%
		
		IF INSTR(WName,PARAM_SET)>0
			FIND_TT+=1
	}
	
	; ----  ITEM FIND
	; ---------------------------------------------------------------
	IF FIND_TT>1
	{
		EXITAPP 1
	}
	
RETURN


