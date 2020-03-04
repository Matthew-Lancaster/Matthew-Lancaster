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
; TO   __ Sun 09-Jun-2019 09:50:00 __ Clipboard Count = 139 __ NEAR 3 HOUR 
; AND THEN 
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------

; -------------------------------------------------------------------
; I DO THIS 
; CHECK ALL SCRIPT THAT HAS LINE IN FILENAME CONTENT OF THE INCLUDE
; CLOSE THEM ALL AND RUN UP AR 
; AND FILTER A FEW OFF
; -------------------------------------------------------------------
; CODER TIME INCLUDE MAKE COOK DINNER LAST BIT ON ABOUT 
; FILENAME EXTRACT STRING WHEN USE SEARCH WINDOW BEFORE
; -------------------------------------------------------------------
; IT CHECK ONLY FOLDER ONE SCRIPT ARE
; AND INCLUDE SUB-FOLDER IF WANT AR
; Loop, Files, %A_ScriptDir%\*.AHK    ; , D
; OR MULTI
; -------------------------------------------------------------------
; I AM IN THE HABIT OF MAKE ARRAY IN A FUNCTION AND PASS THE INFO OVER
; THAT HELP TO MAKE AN ARRAY GLOBAL USE ANYWHERE
; -------------------------------------------------------------------
; SUPER SCRIPT -- INDUSTRIAL SCRIPT AR
; TAKE FOR GRANTER ONCE DONE -- USEFUL TOOL
; -------------------------------------------------------------------
; Sun 26-Jan-2020 12:00:44
; Sun 26-Jan-2020 13:58:00 -- 1 HOUR 57 MINUTE
; -------------------------------------------------------------------
; Sun 26-Jan-2020 13:58:00 -- COOKER
; Sun 26-Jan-2020 16:00:00 -- 2 HOUR 2 MINUTE
; -------------------------------------------------------------------
; Sun 26-Jan-2020 16:00:00 -- ANOTHER BURST CODE -- SORT ICON AHK REMOVE AND RESTORE
; Sun 26-Jan-2020 18:40:00 -- 2 HOUR 40 MINUTE
; -------------------------------------------------------------------
; Sun 26-Jan-2020 12:00:44 -- TOTAL TODAY
; Sun 26-Jan-2020 18:40:00 -- 6 HOUR 39 MINUTE
; -------------------------------------------------------------------

; ---------------------------------------------------------------------------------
; SESSION 002
; ---------------------------------------------------------------------------------
; ROUTINE SET
; ---------------------------------------------------------------------------------
; TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME:
; ---------------------------------------------------------------------------------
; IF KILL OWN PROCESS LAST OF ALL -- IT NOT ABLE REMOVE ICON AND EXIT GRACEFUL THEN
; ---------------------------------------------------------------------------------
; -- Fri 28-Feb-2020 17:24:17
; -- Fri 28-Feb-2020 19:04:28 -- 1 HOUR 40 MINUTE
; ---------------------------------------------------------------------------------



FILE_ScriptName=%A_ScriptName%
IF INSTR(FILE_ScriptName,"_INCLUDE")>0
{
	DetectHiddenWindows, ON
	#SingleInstance force

	; Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk
	FN_ARRAY_INCLUDE_SCRIPT_NAME := ARRAY_INCLUDE_SCRIPT_NAME()
	 
	FILELIST=
	Loop % FN_ARRAY_INCLUDE_SCRIPT_NAME.MaxIndex()
	{
		ELEMENT_22 := FN_ARRAY_INCLUDE_SCRIPT_NAME[A_Index]
		TEMP_VAR_1=%ELEMENT_22% - AutoHotkey v%A_AhkVersion%
		TEMP_VAR_2:=StrReplace(TEMP_VAR_1, """" , "")
		ELEMENT_22=%TEMP_VAR_2%
		IFWINEXIST %ELEMENT_22%
			IF !FILELIST
				FILELIST = %ELEMENT_22%   ; AVOID A BLANK 1ST LINE
			ELSE
				FILELIST = %FILELIST%`n%ELEMENT_22%
	}
	GOSUB PROCESS_KILL_AUTOHOTKEY
	GOSUB RUN_TIMER_TRAY_ICON_CLEAN_UP

	Loop, parse, FILELIST, `n
	{
		REPLACE_STR_FNAME__=- AutoHotkey v%A_AhkVersion%
		FILE_NAME_PATH_____:=StrReplace(A_LoopField,REPLACE_STR_FNAME__,"")
		ifExist %FILE_NAME_PATH_____%
		{
			Run, %FILE_NAME_PATH_____%
			WINWAIT %A_LoopField%
			Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}	

	GOSUB RUN_TIMER_TRAY_ICON_CLEAN_UP
	
EXITAPP
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------

PROCESS_KILL_AUTOHOTKEY:
	SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")

	WinGet, List, List, ahk_exe AutoHotkey.EXE
	PROCESS_COUNT=0
	Loop %List%  
	{
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		If PID_8 <> %SCRIPTOR_OWN_PID%
			PROCESS_COUNT +=1
	}	
	WHILE PROCESS_COUNT>0
	{
			
		Loop %List%  
		{ 
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			If PID_8 <> %SCRIPTOR_OWN_PID%
			{
				Process, Close, %PID_8% 
				Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
			}
		}
		WinGet, List, List, ahk_exe AutoHotkey.EXE
		PROCESS_COUNT=0
		Loop %List%  
		{
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			If PID_8 <> %SCRIPTOR_OWN_PID%
				PROCESS_COUNT +=1
		}	
	}	
RETURN



CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK:
	DetectHiddenWindows, ON
	WinGet, List, List, ahk_class ThunderRT6FormDC
	PATH_ID_BUILD=
	Loop %List%  
	{
		; IfWinExist ahk_id List%A_Index%
		;WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		WinGet, PATH_FULL, ProcessPath, % "ahk_id " List%A_Index% 
		WinGet, PATH_EXE, ProcessName, % "ahk_id " List%A_Index% 
		StringUpper PATH_EXE, PATH_EXE
		StringUpper PATH_FULL, PATH_FULL
			IF INSTR(PATH_FULL,"D:\VB")>0
			{
			
				IF INSTR(PATH_ID_BUILD,PATH_EXE)=0
				{
					PATH_ID_BUILD=%PATH_ID_BUILD%%PATH_EXE%`n
					PROCESS, EXIST, %PATH_EXE%
					IF ERRORLEVEL
					{
						TOOLTIP % PATH_ID_BUILD
						SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
						WINCLOSE ahk_exe %PATH_EXE%
						SLEEP 200
						PROCESS, EXIST, %PATH_EXE%
						IF ERRORLEVEL
						{
							SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							WINCLOSE ahk_exe %PATH_EXE%
						}
						SLEEP 200
						PROCESS, EXIST, %PATH_EXE%
						IF ERRORLEVEL
						{
							SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							Process, Close, %PATH_EXE%
						}
					}
				}

			}
	}	
	TOOLTIP
RETURN



ARRAY_INCLUDE_SCRIPT_NAME() {
	ARRAY_INCLUDE_SCRIPT_NAME := []
	ArrCnt := 0
	Loop, Files, %A_ScriptDir%\*.AHK    ; , D
	{
		StringUpper LoopFileName_UPPER, A_LoopFileName
		IF InStr(LoopFileName_UPPER, "_INCLUDE")=0
		{
			TEMP_VAR_1_FUNCTION=%A_LoopFileName% - AutoHotkey v%A_AhkVersion%
			TEMP_VAR_2_FUNCTION:=StrReplace(TEMP_VAR_1_FUNCTION, """" , "")
			ELEMENT_24=%A_ScriptDir%\%TEMP_VAR_2_FUNCTION%
			
			IFWINEXIST %ELEMENT_24%
			{
				Loop, read, %A_LoopFileName%
				{
					; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk
					; -------------------------------------------------------------------------------------------
					IF InStr(A_LoopReadLine, "#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY")=1
					IF InStr(A_LoopReadLine, "\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk")

						IFWINEXIST %ELEMENT_24%
						{
						FILE_SCRIPT_PATH=%A_ScriptDir%\%A_LoopFileName%
						ArrCnt += 1
						ARRAY_INCLUDE_SCRIPT_NAME[ArrCnt]:=FILE_SCRIPT_PATH
						Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
						BREAK
					}
					IF InStr(LoopFileName_UPPER, "_INCLUDE")>0
						BREAK 
				}
			}
		}
	}
	RETURN ARRAY_INCLUDE_SCRIPT_NAME
}



~<^#ESC:: GOSUB TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME

~>^F1::
{
	GOSUB SUB_RESTORE_VB_KEEP_RUNNER
	GOSUB SUB_RESTORE_ELITESPY
}
RETURN


TIMEZONE_MINI_GUI_DISPLAY_EXE:
Element_1 := "D:\VB6\VB-NT\00_Best_VB_01\TIMEZONE MINI GUI DISPLAY\TIMEZONE MINI GUI DISPLAY.exe"
IfExist, %Element_1%
{
	SoundBeep , 2000 , 100
	Run, %Element_1%
}
RETURN

ALL_LOW_PROCCES_PRIORITY_TO_NORMAL:
Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk"
IfExist, %Element_1%
{
	SoundBeep , 2000 , 100
	Run, %Element_1%
}
RETURN


SET_GITHUB_GO_AR:


	FILE_NET_C=\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_TIMER_GITHUB_CLICKER_#NFS_EX__.txt
	FILE_NET_D=\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_TIMER_GITHUB_CLICKER_#NFS_EX__.txt
	
	FILE_HDD_C=C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_TIMER_GITHUB_CLICKER_#NFS_EX__.txt
	FILE_HDD_D=D:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_TIMER_GITHUB_CLICKER_#NFS_EX__.txt
	
	IF (A_ComputerName="7-ASUS-GL522VW")
	{
		IfNotExist %FILE_HDD_C%
		FileAppend,"INFO_WRITE_FILE",%FILE_HDD_C%
		IfNotExist %FILE_HDD_C%
			FileAppend,"INFO_WRITE_FILE",%FILE_HDD_D%
	}
 
	IfNotExist %FILE_NET_C%
	FileAppend,"INFO_WRITE_FILE",%FILE_NET_C%
	; ---------------------------------------------------------------
	; C DRIVE NOT ACCESS RIGHT HALF THE TIME TRY HDD-D
	; ---------------------------------------------------------------
	IfNotExist %FILE_NET_C%
		FileAppend,"INFO_WRITE_FILE",%FILE_NET_D%

	
	IF (A_ComputerName="7-ASUS-GL522VW")
	{
		IfNotExist %FILE_HDD_C%
		IfNotExist %FILE_HDD_D%
		MSGBOX % "NOT EXIST BOTH `n" FILE_HDD_C "`n`n" FILE_HDD_D
	}
	IF (A_ComputerName<>"7-ASUS-GL522VW")
	{
		IfNotExist %FILE_NET_C%
		IfNotExist %FILE_NET_D%
		MSGBOX % "NOT EXIST BOTH `n" FILE_NET_C "`n`n" FILE_NET_D
	}
	
	; FILE PICKER HERE
	; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
	
RETURN




SUB_RESTORE_VB_KEEP_RUNNER:
	SetTitleMatchMode 2  ; Avoids Specify Full path.
	; WinRestore, VB_KEEP_RUNNER ahk_class ThunderFormDC
	; WinRestore, EliteSpy ahk_class ThunderRT6FormDC
	; WinRestore, VB_KEEP_RUNNER 
	; WinRestore, EliteSpy 
	VB_KEEP_RUNNER_TITLE=VB_KEEP_RUNNER
	IfWinExist, %VB_KEEP_RUNNER_TITLE%
	{
		WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=-1
		{
			WinRestore, %VB_KEEP_RUNNER_TITLE%
			WinActivate, %VB_KEEP_RUNNER_TITLE%
		}
		
		LOOP 200
		{
			WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
			WinGet Style_VB, MinMax, ahk_id %HWND_10%
			IF Style_VB=0
			{
				SOUNDBEEP 1400,100
				BREAK
			}
			IF A_INDEX>150 
				IF Style_VB<>0
				{
					SOUNDBEEP 1400,100
					Process, Close, VB_KEEP_RUNNER.exe
				}
			SLEEP 20	
		}
	}
	IfWinNOTExist, %VB_KEEP_RUNNER_TITLE%
	{
		Process, Close, VB_KEEP_RUNNER.exe
		
		FN_VAR_1 := "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
		IfExist, %FN_VAR_1%
			Run, %FN_VAR_1%
		
		SOUNDBEEP 1400,100
	}
RETURN

SUB_RESTORE_ELITESPY:
	SetTitleMatchMode 2  ; Avoids Specify Full path.
	; WinRestore, VB_KEEP_RUNNER ahk_class ThunderFormDC
	; WinRestore, EliteSpy ahk_class ThunderRT6FormDC
	; WinRestore, VB_KEEP_RUNNER 
	; WinRestore, EliteSpy 
	ELITE_SPY_TITLE=EliteSpy`+ 2001 __ www.PlanetSourceCode.com __ Version
	IfWinExist, %ELITE_SPY_TITLE%
	{
		WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=-1
		{
			WinRestore, %ELITE_SPY_TITLE%
			WinActivate, %ELITE_SPY_TITLE%
		}
		
		LOOP 200
		{
			WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
			WinGet Style_VB, MinMax, ahk_id %HWND_10%
			IF Style_VB=0
			{
				SOUNDBEEP 1400,100
				BREAK
			}
			IF A_INDEX>150 
				IF Style_VB<>0
				{
					SOUNDBEEP 1400,100
					Process, Close, EliteSpy.exe
				}
			SLEEP 20	
		}
	}
	IfWinNOTExist, %ELITE_SPY_TITLE%
	{
		Process, Close, EliteSpy.exe
		
		FN_VAR_1:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
		IfExist, %FN_VAR_1%
			Run, %FN_VAR_1%
		
		SOUNDBEEP 1400,100
	}
RETURN

SUB_RESTORE_VB_KEEP_RUNNER_02:

	; 
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; 
	; THIS IS MY CODER TO BRING UP WINDOW 
	; VB_KEEP_RUNNER 
	; AND
	; ELITE SPY 
	; BUT FOR NOW THESE TWO PROGRAM WHEN COMPUTER UNDER PRESSURE FOR LONG WHEN GET BACK TO
	; THEY BOTH NEVER SEEM RUNNER
	; SO SOMEHOW GOT TO FIND THAT WINDOW POP UP NOT HAPPEN AS NORMAL
	; AND THEN TAKE ACTION TO KILL AND RELOAD
	; I TRY AND CATCH ERROR OF WINDOW-STATE STYLE
	; AND RESULT WAS 0
	; 0 FOR NONE OF MIN OR MAX AS IF WAS NORMAL
	; BUT NOT TRUE 0 WAS RESULT BUT NOT SHOW-ER
	; SO NEXT TIME DEBUG CHECK FOR HIDDEN STATE OR MORE
	; FROM       Sat 11-May-2019 06:04:11
	; STOP TIME  Sat 11-May-2019 07:07:55 -- CODE SOMETHING ELSE -- VB_KEEP_RUNNER -- IMPROVE
	; RESUME-A   Sat 11-May-2019 08:24:53 --
	; TO         Sat 11-May-2019 09:41:33
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	VB_KEEP_RUNNER_VAR=FALSE
	VB_KEEP_RUNNER_TITLE=VB_KEEP_RUNNER
	GetKeyState, state, Shift
	if state = D
	IfWinExist, %VB_KEEP_RUNNER_TITLE%
	{
		WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=-1
		{
			WinRestore, %VB_KEEP_RUNNER_TITLE%
			WinActivate, %VB_KEEP_RUNNER_TITLE%
			SoundBeep , 1000 , 100
			SoundBeep , 2000 , 100
			SoundBeep , 3000 , 100
			VAR_DONE_ESCAPE_KEY=TRUE
			SLEEP 400
		}
		WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=0
		{
			VAR_DONE_ESCAPE_KEY=TRUE
			; VB_KEEP_RUNNER_VAR=TRUE
			; VB_KEEP_RUNNER_VAR_2=TRUE
			; SLEEP 400
		}
	}
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	SetTitleMatchMode 2
	VB_ELITE_SPY_VAR=FALSE
	ELITE_SPY_TITLE=EliteSpy`+ 2001 __ www.PlanetSourceCode.com __ Version
	GetKeyState, state, Shift
	if state = D
	IfWinExist, %ELITE_SPY_TITLE%
	{
		WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=-1
		{
			WinRestore, %ELITE_SPY_TITLE%
	 		WinActivate, %ELITE_SPY_TITLE%
			SoundBeep , 1000 , 100
			SoundBeep , 2000 , 100
			SoundBeep , 3000 , 100
			VAR_DONE_ESCAPE_KEY=TRUE
			VB_ELITE_SPY_VAR=TRUE
		}
		WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		WinGet Style_VB, MinMax, ahk_id %HWND_10%
		IF Style_VB=0
		{
			VAR_DONE_ESCAPE_KEY=TRUE
			; VB_ELITE_SPY_VAR=TRUE
			; VB_ELITE_SPY_VAR_2=TRUE
			; SLEEP 400
		}
	}
	
	VB_KEEP_RUNNER_VAR=FALSE
	VB_KEEP_RUNNER_VAR_2=FALSE
	IF VB_KEEP_RUNNER_VAR=TRUE
	{
		LOOP 
		{
			X_COUNTER+=1
			WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
			WinGet Style_VB, MinMax, ahk_id %HWND_10%
			; IF Style_VB=0
			; MSGBOX % Style_VB
			; MinMax
			; Retrieves the minimized/maximized state for a window.
			; WinGet, OutputVar, MinMax [, WinTitle, WinText, ExcludeTitle, ExcludeText]
			; OutputVar is made blank if no matching window exists; otherwise, it is set to one of the following numbers:
			; •-1: The window is minimized (WinRestore can unminimize it).
			; •1: The window is maximized (WinRestore can unmaximize it).
			; •0: The window is neither minimized nor maximized.


			IF Style_VB=-1
			{
				VB_KEEP_RUNNER_VAR_2=TRUE
				BREAK
			}
			SLEEP 100
			IF X_COUNTER>100
				BREAK
		}
	}
	IfWinNotExist %VB_KEEP_RUNNER_TITLE%
	{
		VAR_DONE_ESCAPE_KEY=TRUE
		SoundBeep , 3000 , 100
		FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
		IfExist, %FN_VAR%
			{
				Run, %FN_VAR% MAXIMUM
			}
	}	
	
	IF VB_KEEP_RUNNER_VAR_2=TRUE
	{
		VAR_DONE_ESCAPE_KEY=TRUE
		WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		WinGet, UniquePID, PID,  ahk_id %HWND_10%
		IF UniquePID>0
		{
			SOUNDBEEP 1000,100
			Process, Close, %UniquePID% 
		}

		IfWinNotExist %VB_KEEP_RUNNER_TITLE%
		{
			SoundBeep , 3000 , 100
			FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
			IfExist, %FN_VAR%
				{
					Run, %FN_VAR% MAXIMUM
				}
		}	
	}
	
	VB_ELITE_SPY_VAR=FALSE
	VB_ELITE_SPY_VAR_2=FALSE
	IF VB_ELITE_SPY_VAR=TRUE
	{
		LOOP 
		{
			X_COUNTER+=1
			WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
			WinGet Style_VB, MinMax, ahk_id %HWND_10%
			; IF Style_VB=0
			; MSGBOX % Style_VB

			IF Style_VB=0
			{
				VB_ELITE_SPY_VAR_2=TRUE
				BREAK
			}
			SLEEP 100
			IF X_COUNTER>100
				BREAK
		}
	}
	IfWinNotExist %ELITE_SPY_TITLE%
	{
		VAR_DONE_ESCAPE_KEY=TRUE
		SoundBeep , 3000 , 100
		FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
		IfExist, %FN_VAR%
			{
				Run, %FN_VAR% MAXIMUM
			}
	}	
	
	IF VB_ELITE_SPY_VAR_2=TRUE
	{
		VAR_DONE_ESCAPE_KEY=TRUE
		WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		WinGet, UniquePID, PID,  ahk_id %HWND_10%
		IF UniquePID>0
		{
			SOUNDBEEP 1000,100
			Process, Close, %UniquePID% 
		}

		IfWinNotExist %ELITE_SPY_TITLE%
		{
			SoundBeep , 3000 , 100
			FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
			IfExist, %FN_VAR%
				{
					Run, %FN_VAR% MAXIMUM
				}
		}	
	}
	
	; ---------------------------------------------------------------
	; # Win (Windows logo key) 
	; ! Alt 
	; ^ Control 
	; + Shift 
	; & An ampersand 
	; ---------------------------------------------------------------


RETURN


; # 
; Win (Windows logo key).
; [v1.0.48.01+]: For Windows Vista and later, hotkeys that include the Win key (e.g. #a) will wait for the Win key to be released before sending any text containing an L keystroke. This prevents usages of Send within such a hotkey from locking the PC. This behavior applies to all sending modes except SendPlay (which doesn't need it) and blind mode. [v1.1.29+]: Text mode is also excluded.

; Note: Pressing a hotkey which includes the Win key may result in extra simulated keystrokes (Ctrl by default). See #MenuMaskKey.
 
; ! Alt
; Note: Pressing a hotkey which includes the Alt key may result in extra simulated keystrokes (Ctrl by default). See #MenuMaskKey.
 
; ^ Control 
; + Shift 


TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME_OLD_BY_LEAVE_SCRIPT_TO_LAST:
TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME:
	DetectHiddenWindows, ON
	SOUNDBEEP 1000,100
	; ---------------------------------------------------------------
	; IN MY CODE THERE IS A PROBLEM 
	; Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
	; THE CODE ABOVE HAS A RELOAD SYSTEM
	; UNLESS IT KILL QUICKER AND WITH REOCCUR
	; IT PRODUCE RELOAD SOME CODE
	; QUICKER AND WITH REOCCUR IS ANSWER SHOVE IN EXTRA CODE
	; ---------------------------------------------------------------

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
	; TO   __ Sun 09-Jun-2019 09:50:00 __ Clipboard Count = 139 __ NEAR 3 HOUR
	; ---------------------------------------------------------------
	
	DetectHiddenWindows, On
	SetTitleMatchMode, 2

	; RUN HERE AS ANOTHER TYPE SCRIPT PROGRAMMER LANGUAGE
	; ----------------------------------------------------
	FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk"
	IfExist, %FN_VAR_1%
	{
		Run, %FN_VAR_1%
		WINWAIT Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk ahk_class AutoHotkey,,30
	}
	
	ScriptName = %FN_VAR_1%
	SCRIPTNAME = %SCRIPTNAME% - AutoHotkey v %A_AhkVersion%
	TT=
	WinGet, List, List, ahk_class AutoHotkey
	Loop %List%
	{ 
		ID_SX:=List%A_Index%
		WinGetTitle, TITLE_VAR_T, ahk_id %ID_SX%
		IF TITLE_VAR_T
		IF INSTR(TITLE_VAR_T,"Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk")>0
			WinGet, PID_78_TRAY_ICON_CLEANER, PID, ahk_id %ID_SX%
		TT=%TT%`n%TITLE_VAR_T%
	}
	
	SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")

	WinGet, List, List, ahk_class AutoHotkey
	Loop
	{
		Loop %List%
		{
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			IF PID_8
			If PID_8 <> %SCRIPTOR_OWN_PID%
			IF PID_8 <> %PID_78_TRAY_ICON_CLEANER%
			{
				PROCESS, Close, %PID_8% 
				SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
				SOUNDBEEP 1000,20
			}
		}
		WinGet, List, List, ahk_class AutoHotkey
		Loop %List%
			I_COUNT=%A_Index%
		SLEEP 50
		IF I_COUNT<3 THEN 
			BREAK
	}
	GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK
	GOSUB KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT
	
	; IF KILL OWN PROCESS LAST OF ALL -- IT NOT ABLE REMOVE ICON AND EXIT GRACEFUL THEN
	; ---------------------------------------------------------------------------------
	; -- Fri 28-Feb-2020 17:24:17
	; -- Fri 28-Feb-2020 19:04:28 -- 1 HOUR 40 MINUTE
	; ---------------------------------------------------------------------------------
	EXITAPP
	RETURN
		
RETURN







KILL_ALL_PROCESS_BY_NAME_CMD_CONHOST_WSCRIPT:
	WinGet, List, List, ahk_exe CMD.EXE
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	WinGet, List, List, ahk_exe CONHOST.EXE
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	WinGet, List, List, ahk_exe WSCRIPT.EXE	
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	
RETURN



TIMER_SUB_AUTOHOTKEY_RELOAD_03:

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
TEMP_VAR_1_INCLUDE=%FN_VAR_1%
TEMP_VAR_2_INCLUDE="%AHK_TERMINATOR_VERSION%"
TEMP_VAR_3=%TEMP_VAR_1_INCLUDE%%TEMP_VAR_2_INCLUDE%
TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")

FN_VAR_2=%TEMP_VAR_3%

IFWINEXIST %FN_VAR_2%
{
	WinGet, PID_01, PID, %FN_VAR_2% ahk_class AutoHotkey
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	SoundBeep , 8000 , 100
	Process, Close,% PID_01
}

IFWINNOTEXIST %FN_VAR_2%
{
	SoundBeep , 5000 , 100
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, %TEMP_VAR_1_INCLUDE%
}

RETURN




RELOAD_ALL_NET___VB_CODE_EXE_SUB:

		FILENAME_2__=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE
		
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
				ELEMENT5=%FILENAME_2__%
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
		
		FILENAME_2__=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE
		
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
				ELEMENT5=%FILENAME_2__%
				NET_PATH:=A_LoopReadLine
				ELEMENT7=_%NET_PATH%.TXT

				Array_FileName%ArrayCount% =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
			}
		}

		Loop %ArrayCount%
		{
			file := FileOpen(Array_FileName%A_Index%, "w")
			if !IsObject(file)
			{
				MsgBox Can't open "%FileName%" for writing.
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

	FILENAME_2__=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NETWORK_VB_CODE_EXE_%A_ComputerName%.TXT
	
	; SCRIPT_DIR__:=SubStr(A_ScriptDir, 3)       ; FROM 3 ONWARD INCLUDE 1ST \
	; StringReplace, SCRIPT_DIR__,SCRIPT_DIR__,\SCRIPTER\,\SCRIPTOR DATA\, All

	StringReplace, FILENAME_2__,FILENAME_2__,\SCRIPTER\,\SCRIPTOR DATA\, All

	
	NET_PATH = %A_ComputerName%

	ELEMENT1=\\
	ELEMENT2=%NET_PATH%
	NET_PATH := StrReplace(NET_PATH, "-", "_")
	ELEMENT3=\
	ELEMENT4=%NET_PATH%
	ELEMENT5=%FILENAME_2__%

	FILENAME_2__ =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%
	
	FileExist_FLAG=FALSE
	if FileExist(FILENAME_2__)
		FileExist_FLAG=TRUE
		
	; IF FileExist_FLAG<>%OLD_FileExist_FLAG%
		IF FileExist(FILENAME_2__)
		{
		
			WINCLOSE EliteSpy+ by Andrea B 2001 __
			WINCLOSE VB_KEEP_RUNNER
			WINCLOSE INDIVIDUAL PROCESS _ Ver
			
			; Process, Close, VB_KEEP_RUNNER.exe
		}		

	; IF FileExist_FLAG<>%OLD_FileExist_FLAG%
		IF !FileExist(FILENAME_2__)
		{
		
			GOSUB TIMER_SUB_EliteSpy
			GOSUB TIMER_SUB_VB_KEEP_RUNNER
			; GOSUB TIMER_SUB_CPU_INDIVIDUAL_PROCESS
			
		}		
		
	OLD_FileExist_FLAG=%FileExist_FLAG%

DetectHiddenWindows, % dhw
	
RETURN



---------------------------


; 01 OF 04 ---- RESTORE CREATE FILE DELETE AND RELOAD
RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_01_OF_04:

	OPERATION_CREATE_PATH_SET_NETWORK=VB
	GOSUB RESTORE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
		
	GOSUB TIMER_SUB_EliteSpy
	GOSUB TIMER_SUB_VB_KEEP_RUNNER
	; GOSUB TIMER_SUB_CPU_INDIVIDUAL_PROCESS
		
	FileDelete, % FileName_4
	
RETURN

; 02 OF 04 ---- DELETE AND KILL PROCESS
RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_02_OF_04:
	
	OPERATION_CREATE_PATH_SET_NETWORK=VB
	GOSUB DELETE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------

	WINCLOSE EliteSpy+ by Andrea B 2001 __
	WINCLOSE VB_KEEP_RUNNER
	WINCLOSE INDIVIDUAL PROCESS _ Ver
		
	; Process, Close, VB_KEEP_RUNNER.exe

	;FileName_RESTORE_VB_AHK=%FileName_4%
	;GOSUB CREATE_FILENAME_FORMAT_LOCAL_LEVEL_FROM_PARAM

RETURN

; 03 OF 04 ---- RESTORE CREATE FILE DELETE AND RELOAD
RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_03_OF_04:

	OPERATION_CREATE_PATH_SET_NETWORK=AHK
	GOSUB RESTORE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------

	GOSUB TIMER_SUB_AUTOHOTKEY_RELOAD_03

	GOSUB RETURN_FILENAME_FORMAT_LOCAL_LEVEL_VB_AND_AHK
	FileDelete, % FileName_AHK
	
RETURN


; 04 OF 04 ---- DELETE AND KILL PROCESS
RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_04_OF_04:
	
	OPERATION_CREATE_PATH_SET_NETWORK=AHK
	GOSUB DELETE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE

	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	
	SetTitleMatchMode 2  ; Avoid specify the full path of the file below.
	DetectHiddenWindows, On 
	; -----------------------------------------------------------
	; THIS ROUTINE WILL KILL ALL AUTOHOTKEYS IN ONE LINE BUT 
	; EXTERNAL CODE OF VBSCRIPT
	; DOESN'T MATTER ABOUT KILL ITSELF LAST THEN
	; FOLLOWER ON IS OTHER EXAMPLE TO DO IN AUTOHOTKEYS
	; THIS CODE JUMP TO VBSCRIPT WHICH JUMP TO BAT LANGUAGE ALSO
	; -----------------------------------------------------------
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS"
	; FN_VAR:="" ; TEMP AVOID RATHER HAVE NONE DISPLAY
	IFWINNOTEXIST, SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%"  /F /IM AutoHotKey.exe /T , , Max
		}
	
	; NEXT BIT OF CODE WILL RUN BEFORE TASKKILL DONE DISPLAY OF WORKER UNLESS DELAY
	; --------------------------------------------
	SLEEP 2000
	
	; -----------------------------------------------------------
	; THIS ROUTINE CODE WILL KILL ALL AHK 
	; BUT NOT INCLUDING MESSAGE BOX ITEM
	; -----------------------------------------------------------
	WinGet, List, List, ahk_class AutoHotkey 
	Loop %List% 
	  { 
		WinGet, PID_02, PID, % "ahk_id " List%A_Index% 
		If ( PID_02 <> DllCall("GetCurrentProcessId") )
			{
				Process, Close, % PID_02
				; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				SoundBeep , 2500 , 100
			}
	  }
	
	; -----------------------------------------------------------
	; THIS ROUTINE CODE WILL KILL ALL AHK 
	; INCLUDING MSGBOX MESSAGE BOX ITEM
	; -----------------------------------------------------------
	PID_02:=DllCall("GetCurrentProcessId") 
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where name = 'Autohotkey.exe' and processID  <> " PID_02 )
	{
		Process, Close, % process.ProcessId
		SoundBeep , 2500 , 100
	}
	; -----------------------------------------------------------
	; KILL OUR OWN SCRIPTOR - LASTLY
	; -----------------------------------------------------------
	Process, Close,% DllCall("GetCurrentProcessId")
	SoundBeep , 2500 , 100
	WINCLOSE % A_ScriptName
	SoundBeep , 2500 , 100

	; CREDIT SOURCE KICK START
	; ----
	; a simple script to kill all autohotkey.exe ? - Ask for Help - AutoHotkey Community
	; https://autohotkey.com/board/topic/75424-a-simple-script-to-kill-all-autohotkeyexe/
	; ----
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
			
	;FileName_RESTORE_VB_AHK=%FileName_4%
	;GOSUB CREATE_FILENAME_FORMAT_LOCAL_LEVEL_FROM_PARAM
RETURN

RESTORE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE:
	GOSUB CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE
	GOSUB CREATE_FILENAME_FORMAT_ALL_NETWORK_LEVEL_FROM_ARRAY
RETURN

DELETE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE:
	GOSUB CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE
	Loop %ArrayCount%
	{
		FileDelete, % Array_FILENAME_2__[A_Index]
		SOUNDBEEP 1000,100
	}
RETURN

CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE:

	FileName_O_VB=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE
	FileName_O_AHK=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE

	IF OPERATION_CREATE_PATH_SET_NETWORK=VB
	{
		FILENAME_2__=%FileName_O_VB%
		FILENAME_4:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE_%A_ComputerName%.TXT"
	}
	IF OPERATION_CREATE_PATH_SET_NETWORK=AHK
	{
		FILENAME_2__=%FileName_O_AHK%	
		FILENAME_4:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE_%A_ComputerName%.TXT"
	}
	
	GOSUB READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY
	
RETURN

READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY:

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
			ELEMENT5=%FILENAME_2__%
			NET_PATH:=A_LoopReadLine
			ELEMENT7=_%NET_PATH%.TXT

			Array_FileName%ArrayCount% =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
			Array_FILENAME_2__[ArrayCount] := Array_FileName%ArrayCount%
			; MSGBOX % Array_FILENAME_2__[ArrayCount]
		}
	}

RETURN

CREATE_FILENAME_FORMAT_LOCAL_LEVEL_FROM_PARAM:

	FileName := FileOpen(FileName_4, "w")
	if !IsObject(FileName)
	{
		MsgBox Can't open "%FileName_4%" for writing.
		return
	}
	TestString := "This is a test string.`r`n"  
	FileName.Write(TestString)
	FileName.Close()

RETURN 

CREATE_FILENAME_FORMAT_ALL_NETWORK_LEVEL_FROM_ARRAY:

	Loop %ArrayCount%
	{
		FileName := FileOpen(Array_FILENAME_2__[A_Index], "w")
		if !IsObject(FileName)
		{
			MsgBox Can't open "%FileName%" for writing.
			return
		}
		TestString := "This is a test string.`r`n"  
		FileName.Write(TestString)
		FileName.Close()
		SOUNDBEEP 1000,100
	}

RETURN


TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY:
	; SETTIMER TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE,OFF
	; SETTIMER TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY, OFF
	TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY_VAR=FALSE 
RETURN

RETURN_FILENAME_FORMAT_LOCAL_LEVEL_VB_AND_AHK:

	FileName_VB =_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE_%A_ComputerName%.TXT
	FileName_AHK=_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE_%A_ComputerName%.TXT

	NET_PATH = %A_ComputerName%
	ELEMENT1=\\
	ELEMENT2=%NET_PATH%
	NET_PATH := StrReplace(NET_PATH, "-", "_")
	ELEMENT3=\
	ELEMENT4=%NET_PATH%
	ELEMENT5=%FileName_VB%
	FileName_VB  =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%
	ELEMENT5=%FileName_AHK%
	FileName_AHK =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%
	
RETURN

TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE:
	
	; JUST IN CASE TIMER DON'T LEAVE RUNNER
	; ---------------------------------------------------------------
	; IF TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY_VAR=FALSE 
	; {
		; ; SETTIMER TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY, 10000
		; TIMER_KILL_RELOAD_ALL_NET_ARRAY_CODE_EXE_STOP_DELAY_VAR=TRUE
	; }
	
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, OFF
	SetTitleMatchMode 2  ; Avoids the need to specify the full path

	GOSUB RETURN_FILENAME_FORMAT_LOCAL_LEVEL_VB_AND_AHK
	
	; ---------------------------------------------------------------
	; RUN ONCE CODE AT START SCRIPT
	; AND THEN TOGGLE DON'T RUN BEGIN
	; CHANGE DETECTOR
	; ---------------------------------------------------------------
	IF FileExist_FLAG_RUN_IN_ONCE_FLAG=FALSE 
	{
		FileExist_FLAG_01=FALSE
		if FileExist(FileName_VB)
			FileExist_FLAG_01=TRUE
		FileExist_FLAG_02=FALSE
		if FileExist(FileName_AHK)
			FileExist_FLAG_02=TRUE

		OLD_FileExist_FLAG_01=%FileExist_FLAG_01%
		OLD_FileExist_FLAG_02=%FileExist_FLAG_02%
		FileExist_FLAG_RUN_IN_ONCE_FLAG=TRUE
	}
	
	FileExist_FLAG_01=FALSE
	if FileExist(FileName_VB)
		FileExist_FLAG_01=TRUE
	FileExist_FLAG_02=FALSE
	if FileExist(FileName_AHK)
		FileExist_FLAG_02=TRUE
		
	IF FileExist_FLAG_01<>%OLD_FileExist_FLAG_01%
	{
		; 01 OF 04 ---- FileExist YES
		IF FileExist(FileName_VB)
		{
			; 01 OF 04 ---- DELETE AND KILL PROCESS
			GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_01_OF_04
		}		

		; 02 OF 04 ---- FileExist NOT
		IF !FileExist(FileName_VB)
		{
			; 02 OF 04 ---- RESTORE CREATE FILE AND RELOAD
			GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_02_OF_04
		}		
	}

	IF FileExist_FLAG_02<>%OLD_FileExist_FLAG_02%
	{

		; 03 OF 04  ---- FileExist YES
		IF FileExist(FileName_AHK)
		{
			; 03 OF 04 ---- DELETE AND KILL PROCESS
			GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_03_OF_04
		}		

		; 04 OF 04  ---- FileExist NOT
		IF !FileExist(FileName_AHK)
		{
			; 04 OF 04 ---- RESTORE CREATE FILE AND RELOAD
			GOSUB RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_04_OF_04
		}		
	}
	
	OLD_FileExist_FLAG_01=%FileExist_FLAG_01%
	OLD_FileExist_FLAG_02=%FileExist_FLAG_02%

DetectHiddenWindows, % dhw

RETURN



; -------------------------------------------------------------------
; TO DO WITHER -- Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
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
; TO DO WITHER -- Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
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
; TO DO WITHER -- Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
TIMER_SUB_VB_KEEP_RUNNER:
; -------------------------------------------------------------------
dhw := A_DetectHiddenWindows
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
DetectHiddenWindows, % dhw
Return

; -------------------------------------------------------------------
; TO DO WITHER -- Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
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

 


WRITE_FILE_SCREEN_BRIGHT_FOR_1_HOUR:
		
		FILENAME_2__=_01_c_drive\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER ADD_SOME_#NFS_EX_
		
		; -----------------------------------------------------------
		; WORK AS ELEMENT7 BEGIN WITH UNDERLINE DASH -- MAKE IT DOUBLE DASH FROM ABOVE
		; AND MIS-MATCHER
		; Mon 17-Feb-2020 04:58:54
		; ALL WORK DONE FOR DAY
		; ABOUT FIVE OR FOUR WORK TONIGHT 
		; HA HA HA
		; -----------------------------------------------------------
		
		
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
				ELEMENT5=%FILENAME_2__%
				; NET_PATH:=A_LoopReadLine
				ELEMENT7=%NET_PATH%.TXT

				Array_FileName_VAR=%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
				Array_FileName%ArrayCount%=%Array_FileName_VAR%
				; MSGBOX %Array_FileName_VAR%
			}
		}

		Loop %ArrayCount%
		{
			file := FileOpen(Array_FileName%A_Index%, "w")
			if IsObject(file)
			{
				; MSGBOX % Array_FileName%A_Index%

				; if !IsObject(file)
				; MsgBox Can't open "%FileName%" for writing.
				; return

				TestString := "This is a test string.`r`n"  
				file.Write(TestString)
				file.Close()
			}
		}
		SOUNDBEEP 2000,100

RETURN





RUN_TIMER_TRAY_ICON_CLEAN_UP:

	Array_Icon_GetInfo := TrayIcon_GetInfo_MENU_03()
	Loop % Array_Icon_GetInfo.MaxIndex()
	{

		IF Array_Icon_GetInfo[A_Index].process<1
		{
			TrayIcon_Remove_MENU_03(Array_Icon_GetInfo[A_Index].HWND, Array_Icon_GetInfo[A_Index].uID)
			Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		}
	}
	
	; ---------------------------------------------------------------
	; THIS THE SOFTWARE TO DEVICE DRIVER THE CSR BLUETOOTH 4.0 BY STAR-TECH
	; I GOT TWO COMPUTER SAME BUT ONE DOES NOT RUN UPDATES FOR WINDOWS 10
	; AND SO THE SOFTWARE HAS BE INSTALLED AGAINST ADVISE AS WINDOWS 10 NORMAL 
	; INSTALL DEVICE WON'T WORK
	; GOT EVERYTHING GOING NORMAL AS RESULT THIS
	; BUT LEAVE BEHIND AN ICON IN TASK TRAY UNABLE REMOVE ANY OTHER WAY
	; ---------------------------------------------------------------
	; C:\Program Files\CSR\CSR Harmony Wireless Software Stack\TrayApplication.exe
	; ---------------------------------------------------------------
	; Loop % Array_Icon_GetInfo.MaxIndex()
	; {
		; IF Array_Icon_GetInfo[A_Index].process="TrayApplication.exe"
			; TrayIcon_Remove_MENU_03(Array_Icon_GetInfo[A_Index].HWND, Array_Icon_GetInfo[A_Index].uID)
	; }
	
RETURN
	

; ----------------------------------------------------------------------------------------------------------------------
; Name ..........: TrayIcon library
; Description ...: Provide some useful functions to deal with Tray icons.
; AHK Version ...: AHK_L 1.1.22.02 x32/64 Unicode
; Original Author: Sean (http://goo.gl/dh0xIX) (http://www.autohotkey.com/forum/viewtopic.php?t=17314)
; Update Author .: Cyruz (http://ciroprincipe.info) (http://ahkscript.org/boards/viewtopic.php?f=6&t=1229)
; Mod Author ....: Fanatic Guru
; License .......: WTFPL - http://www.wtfpl.net/txt/copying/
; Version Date...: 2019 04 04
; Note ..........: Many people have updated Sean's original work including me but Cyruz's version seemed the most straight
; ...............: forward update for 64 bit so I adapted it with some of the features from my Fanatic Guru version.
; Update 20160120: Went through all the data types in the DLL and NumGet and matched them up to MSDN which fixed IDcmd.
; Update 20160308: Fix for Windows 10 NotifyIconOverflowWindow
; Update 20180313: Fix problem with "VirtualFreeEx" pointed out by nnnik
; Update 20180313: Additional fix for previous Windows 10 NotifyIconOverflowWindow fix breaking non-hidden icons
; Update 20190404: Added TrayIcon_Set by Cyruz
; ----------------------------------------------------------------------------------------------------------------------

; ----------------------------------------------------------------------------------------------------------------------
; Function ......: TrayIcon_GetInfo
; Description ...: Get a series of useful information about tray icons.
; Parameters ....: sExeName  - The exe for which we are searching the tray icon data. Leave it empty to receive data for 
; ...............:             all tray icons.
; Return ........: oTrayIcon_GetInfo - An array of objects containing tray icons data. Any entry is structured like this:
; ...............:             oTrayIcon_GetInfo[A_Index].idx     - 0 based tray icon index.
; ...............:             oTrayIcon_GetInfo[A_Index].IDcmd   - Command identifier associated with the button.
; ...............:             oTrayIcon_GetInfo[A_Index].pID     - Process ID.
; ...............:             oTrayIcon_GetInfo[A_Index].uID     - Application defined identifier for the icon.
; ...............:             oTrayIcon_GetInfo[A_Index].msgID   - Application defined callback message.
; ...............:             oTrayIcon_GetInfo[A_Index].hIcon   - Handle to the tray icon.
; ...............:             oTrayIcon_GetInfo[A_Index].hWnd    - Window handle.
; ...............:             oTrayIcon_GetInfo[A_Index].Class   - Window class.
; ...............:             oTrayIcon_GetInfo[A_Index].Process - Process executable.
; ...............:             oTrayIcon_GetInfo[A_Index].Tray    - Tray Type (Shell_TrayWnd or NotifyIconOverflowWindow).
; ...............:             oTrayIcon_GetInfo[A_Index].tooltip - Tray icon tooltip.
; Info ..........: TB_BUTTONCOUNT message - http://goo.gl/DVxpsg
; ...............: TB_GETBUTTON message   - http://goo.gl/2oiOsl
; ...............: TBBUTTON structure     - http://goo.gl/EIE21Z
; ----------------------------------------------------------------------------------------------------------------------

TrayIcon_GetInfo_MENU_03(sExeName := "")
{
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	oTrayIcon_GetInfo := {}
	For key, sTray in ["Shell_TrayWnd", "NotifyIconOverflowWindow"]
	{
		idxTB := TrayIcon_GetTrayBar_MENU_03(sTray)
		WinGet, pidTaskbar, PID, ahk_class %sTray%
		
		hProc := DllCall("OpenProcess", UInt, 0x38, Int, 0, UInt, pidTaskbar)
		pRB   := DllCall("VirtualAllocEx", Ptr, hProc, Ptr, 0, UPtr, 20, UInt, 0x1000, UInt, 0x4)

		SendMessage, 0x418, 0, 0, ToolbarWindow32%idxTB%, ahk_class %sTray%   ; TB_BUTTONCOUNT
		
		szBtn := VarSetCapacity(btn, (A_Is64bitOS ? 32 : 20), 0)
		szNfo := VarSetCapacity(nfo, (A_Is64bitOS ? 32 : 24), 0)
		szTip := VarSetCapacity(tip, 128 * 2, 0)
		
		Loop, %ErrorLevel%
		{
			SendMessage, 0x417, A_Index - 1, pRB, ToolbarWindow32%idxTB%, ahk_class %sTray%   ; TB_GETBUTTON
			DllCall("ReadProcessMemory", Ptr, hProc, Ptr, pRB, Ptr, &btn, UPtr, szBtn, UPtr, 0)

			iBitmap := NumGet(btn, 0, "Int")
			IDcmd   := NumGet(btn, 4, "Int")
			statyle := NumGet(btn, 8)
			dwData  := NumGet(btn, (A_Is64bitOS ? 16 : 12))
			iString := NumGet(btn, (A_Is64bitOS ? 24 : 16), "Ptr")

			DllCall("ReadProcessMemory", Ptr, hProc, Ptr, dwData, Ptr, &nfo, UPtr, szNfo, UPtr, 0)

			hWnd_I:= NumGet(nfo, 0, "Ptr")
			uID   := NumGet(nfo, (A_Is64bitOS ? 8 : 4), "UInt")
			msgID := NumGet(nfo, (A_Is64bitOS ? 12 : 8))
			hIcon := NumGet(nfo, (A_Is64bitOS ? 24 : 20), "Ptr")

			WinGet, pID_TrayIcon, PID, ahk_id %hWnd_I%
			WinGet, sProcess, ProcessName, ahk_id %hWnd_I%
			WinGetClass, sClass, ahk_id %hWnd_I%

			If !sExeName || (sExeName = sProcess) || (sExeName = pID_TrayIcon)
			{
				DllCall("ReadProcessMemory", Ptr, hProc, Ptr, iString, Ptr, &tip, UPtr, szTip, UPtr, 0)
				Index := (oTrayIcon_GetInfo.MaxIndex()>0 ? oTrayIcon_GetInfo.MaxIndex()+1 : 1)
				oTrayIcon_GetInfo[Index,"idx"]     := A_Index - 1
				oTrayIcon_GetInfo[Index,"IDcmd"]   := IDcmd
				oTrayIcon_GetInfo[Index,"pID"]     := pID_TrayIcon
				oTrayIcon_GetInfo[Index,"uID"]     := uID
				oTrayIcon_GetInfo[Index,"msgID"]   := msgID
				oTrayIcon_GetInfo[Index,"hIcon"]   := hIcon
				oTrayIcon_GetInfo[Index,"hWnd"]    := hWnd_I
				oTrayIcon_GetInfo[Index,"Class"]   := sClass
				oTrayIcon_GetInfo[Index,"Process"] := sProcess
				oTrayIcon_GetInfo[Index,"Tooltip"] := StrGet(&tip, "UTF-16")
				oTrayIcon_GetInfo[Index,"Tray"]    := sTray
			}
		}
		DllCall("VirtualFreeEx", Ptr, hProc, Ptr, pRB, UPtr, 0, Uint, 0x8000)
		DllCall("CloseHandle", Ptr, hProc)
	}
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
	Return oTrayIcon_GetInfo
}

; -------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Hide
; Description ..: Hide or unhide a tray icon.
; Parameters ...: IDcmd - Command identifier associated with the button.
; ..............: bHide - True for hide, False for unhide.
; ..............: sTray - 1 or Shell_TrayWnd || 0 or NotifyIconOverflowWindow.
; Info .........: TB_HIDEBUTTON message - http://goo.gl/oelsAa
; -------------------------------------------------------------------------------------------------
TrayIcon_Hide_MENU_03(IDcmd, sTray := "Shell_TrayWnd", bHide:=True)
{
	(sTray == 0 ? sTray := "NotifyIconOverflowWindow" : sTray == 1 ? sTray := "Shell_TrayWnd" : )
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	idxTB := TrayIcon_GetTrayBar_MENU_03()
	SendMessage, 0x404, IDcmd, bHide, ToolbarWindow32%idxTB%, ahk_class %sTray% ; TB_HIDEBUTTON
	SendMessage, 0x1A, 0, 0, , ahk_class %sTray%
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Delete
; Description ..: Delete a tray icon.
; Parameters ...: idx - 0 based tray icon index.
; ..............: sTray - 1 or Shell_TrayWnd || 0 or NotifyIconOverflowWindow.
; Info .........: TB_DELETEBUTTON message - http://goo.gl/L0pY4R
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_Delete_MENU_03(idx, sTray := "Shell_TrayWnd")
{
	(sTray == 0 ? sTray := "NotifyIconOverflowWindow" : sTray == 1 ? sTray := "Shell_TrayWnd" : )
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	idxTB := TrayIcon_GetTrayBar_MENU_03()
	SendMessage, 0x416, idx, 0, ToolbarWindow32%idxTB%, ahk_class %sTray% ; TB_DELETEBUTTON
	SendMessage, 0x1A, 0, 0, , ahk_class %sTray%
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Remove
; Description ..: Remove a tray icon.
; Parameters ...: hWnd_I, uID.
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_Remove_MENU_03(hWnd_I, uID)
{
		NumPut(VarSetCapacity(NID,(A_IsUnicode ? 2 : 1) * 384 + A_PtrSize * 5 + 40,0), NID)
		NumPut(hWnd_I , NID, (A_PtrSize == 4 ? 4 : 8 ))
		NumPut(uID    , NID, (A_PtrSize == 4 ? 8  : 16 ))
		Return DllCall("shell32\Shell_NotifyIcon", "Uint", 0x2, "Uint", &NID)
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Move
; Description ..: Move a tray icon.
; Parameters ...: idxOld - 0 based index of the tray icon to move.
; ..............: idxNew - 0 based index where to move the tray icon.
; ..............: sTray - 1 or Shell_TrayWnd || 0 or NotifyIconOverflowWindow.
; Info .........: TB_MOVEBUTTON message - http://goo.gl/1F6wPw
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_Move_MENU_03(idxOld, idxNew, sTray := "Shell_TrayWnd")
{
	(sTray == 0 ? sTray := "NotifyIconOverflowWindow" : sTray == 1 ? sTray := "Shell_TrayWnd" : )
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	idxTB := TrayIcon_GetTrayBar_MENU_03()
	SendMessage, 0x452, idxOld, idxNew, ToolbarWindow32%idxTB%, ahk_class %sTray% ; TB_MOVEBUTTON
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Set
; Description ..: Modify icon with the given index for the given window.
; Parameters ...: hWnd_I     - Window handle.
; ..............: uId        - Application defined identifier for the icon.
; ..............: hIcon      - Handle to the tray icon.
; ..............: hIconSmall - Handle to the small icon, for window menubar. Optional.
; ..............: hIconBig   - Handle to the big icon, for taskbar. Optional.
; Return .......: True on success, false on failure.
; Info .........: NOTIFYICONDATA structure  - https://goo.gl/1Xuw5r
; ..............: Shell_NotifyIcon function - https://goo.gl/tTSSBM
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_Set_MENU_03(hWnd_I, uId, hIcon, hIconSmall:=0, hIconBig:=0)
{
    d := A_DetectHiddenWindows
    DetectHiddenWindows, On
    ; WM_SETICON = 0x0080
    If ( hIconSmall ) 
        SendMessage, 0x0080, 0, hIconSmall,, ahk_id %hWnd_I%
    If ( hIconBig )
        SendMessage, 0x0080, 1, hIconBig,, ahk_id %hWnd_I%
    DetectHiddenWindows, %d%

    VarSetCapacity(NID, szNID := ((A_IsUnicode ? 2 : 1) * 384 + A_PtrSize*5 + 40),0)
    NumPut( szNID,   NID, 0                           )
    NumPut( hWnd_I,  NID, (A_PtrSize == 4) ? 4   : 8  )
    NumPut( uId,     NID, (A_PtrSize == 4) ? 8   : 16 )
    NumPut( 2,       NID, (A_PtrSize == 4) ? 12  : 20 )
    NumPut( hIcon,   NID, (A_PtrSize == 4) ? 20  : 32 )
    
    ; NIM_MODIFY := 0x1
    Return DllCall("Shell32.dll\Shell_NotifyIcon", UInt,0x1, Ptr,&NID)
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_GetTrayBar
; Description ..: Get the tray icon handle.
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_GetTrayBar_MENU_03(Tray:="Shell_TrayWnd")
{
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	WinGet, ControlList, ControlList, ahk_class %Tray%
	RegExMatch(ControlList, "(?<=ToolbarWindow32)\d+(?!.*ToolbarWindow32)", nTB)
	Loop, %nTB%
	{
		ControlGet, hWnd_TrayIcon, hWnd,, ToolbarWindow32%A_Index%, ahk_class %Tray%
		hParent := DllCall( "GetParent", Ptr, hWnd_TrayIcon )
		WinGetClass, sClass, ahk_id %hParent%
		If !(sClass = "SysPager" or sClass = "NotifyIconOverflowWindow" )
			Continue
		idxTB := A_Index
		Break
	}
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
	Return  idxTB
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_GetHotItem
; Description ..: Get the index of tray's hot item.
; Info .........: TB_GETHOTITEM message - http://goo.gl/g70qO2
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_GetHotItem_MENU_03()
{
	idxTB := TrayIcon_GetTrayBar_MENU_03()
	SendMessage, 0x447, 0, 0, ToolbarWindow32%idxTB%, ahk_class Shell_TrayWnd ; TB_GETHOTITEM
	Return ErrorLevel << 32 >> 32
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: TrayIcon_Button
; Description ..: Simulate mouse button click on a tray icon.
; Parameters ...: sExeName - Executable Process Name of tray icon.
; ..............: sButton  - Mouse button to simulate (L, M, R).
; ..............: bDouble  - True to double click, false to single click.
; ..............: index    - Index of tray icon to click if more than one match.
; ----------------------------------------------------------------------------------------------------------------------
TrayIcon_Button_MENU_03(sExeName, sButton := "L", bDouble := false, index := 1)
{
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	WM_MOUSEMOVE	  = 0x0200
	WM_LBUTTONDOWN	  = 0x0201
	WM_LBUTTONUP	  = 0x0202
	WM_LBUTTONDBLCLK = 0x0203
	WM_RBUTTONDOWN	  = 0x0204
	WM_RBUTTONUP	  = 0x0205
	WM_RBUTTONDBLCLK = 0x0206
	WM_MBUTTONDOWN	  = 0x0207
	WM_MBUTTONUP	  = 0x0208
	WM_MBUTTONDBLCLK = 0x0209
	sButton := "WM_" sButton "BUTTON"
	oIcons := {}
	oIcons := TrayIcon_GetInfo_MENU_03(sExeName)
	msgID  := oIcons[index].msgID
	uID    := oIcons[index].uID
	hWnd_I := oIcons[index].hWnd
	if bDouble
		PostMessage, msgID, uID, %sButton%DBLCLK, , ahk_id %hWnd_I%
	else
	{
		PostMessage, msgID, uID, %sButton%DOWN, , ahk_id %hWnd_I%
		PostMessage, msgID, uID, %sButton%UP, , ahk_id %hWnd_I%
	}
	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
	return
}



