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

~<^#ESC:: GOSUB TERMINATE_ALL_AUTOHOTKEYS_SCRIPT_BY_EXE_NAME

~>^F1::
{
	GOSUB SUB_RESTORE_VB_KEEP_RUNNER
	GOSUB SUB_RESTORE_ELITESPY
}
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
	
	SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")

	Allitem=2
	WHILE Allitem>1
	{
		; WinGet, List, List, ahk_exe AutoHotkey.EXE
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List%  
		{ 
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			IF PID_8
			If PID_8 <> %SCRIPTOR_OWN_PID%
			{
				; MSGBOX % PID_8 " -- " SCRIPTOR_OWN_PID
				; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				Process, Close, %PID_8% 
				SOUNDBEEP 1200,40
			}
		}
		Allitem=0
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List%  
			Allitem:=A_Index
	}	
	
	; ---------------------------------------------------------------
	; SOMETIME AUTOHOTKEY IS LEFT BEHIND WHEN LOOK FOR CLASS
	; AS SOME MSGBOX-ER ARE DIFFERENT CLASS NAME
	; ---------------------------------------------------------------
	Allitem=2
	WHILE Allitem>1
	{
		WinGet, List, List, ahk_exe AutoHotkey.EXE
		Loop %List%  
		{ 
			; -------------------------------------------------------
			; MSGBOX % List%A_Index% -- HWND NUMBER IS HERE
			; BUT I WAS THINKER -- THAT STAY WITH HWND NUMBER WOULD OF BEEN BETTER QUICKER
			; WHEN MACHINE UNDER LOAD ABLE TO SEE STRESS MAKE
			; LESS SPEED BETWEEN EACH KILL
			; OH GUESS REQUIRING PROCESS NUMBER AT END ANYHOW
			; WAS THINKER COULD JUST KILL ALL AHK
			; WHICH WOULD NOT BE SUCH A BAD IDEA
			; SOMEBODY ONCE TALKER THE SCRIPT REQUEST LIST LEAVES ITSELF TO LAST ANYWAY
			; WHAT A KNOBBER LOSER IT CHECK PID FOR OWN EACH TIME GET A NEW NUMBER
			; DONE THAT
			; IT NOT THIS BIT THAT LESS QUICKER
			; IT THE REMOVE ICON
			; -------------------------------------------------------
			IF PID_8
			WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
			If PID_8 <> %SCRIPTOR_OWN_PID%
			{
				; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				Process, Close, %PID_8% 
				SOUNDBEEP 1200,40
			}
		}
		Allitem=0
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List%  
			Allitem:=A_Index
	}	

	GOSUB KILL_ALL_PROCESS_BY_NAME
	
	
	
	; FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - RUN_ONCE.ahk"
	; IfExist, %FN_VAR_1%
	    ; Run, %FN_VAR_1%,,HIDE
	
	FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - WAIT RUN_ONCE.ahk"
	IfExist, %FN_VAR_1%
		Run, %FN_VAR_1%

	; Run, %FN_VAR_1%,,HIDE -- HIDE NOT WORKER FOR AHK
	; USE
	; #NoTrayIcon
		
	Process, Close, %SCRIPTOR_OWN_PID%
RETURN







KILL_ALL_PROCESS_BY_NAME:
	WinGet, List, List, ahk_exe cmd.exe
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
		}
	}
	WinGet, List, List, ahk_exe conhost.exe
	Loop %List%  
	{ 
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		IF PID_8
		{
			Process, Close, %PID_8% 
			SOUNDBEEP 1200,40
		}
	}
RETURN



TIMER_SUB_AUTOHOTKEY_RELOAD_03:

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
TEMP_VAR_1=%FN_VAR_1%
TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")

FN_VAR_2=%TEMP_VAR_3%

IFWINEXIST %FN_VAR_2%
{
	WinGet, PID_01, PID, %FN_VAR_2% ahk_class AutoHotkey
	SoundBeep , 8000 , 100
	Process, Close,% PID_01
}

IFWINNOTEXIST %FN_VAR_2%
{
	SoundBeep , 5000 , 100
	Run, %TEMP_VAR_1%
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
		
		FILENAME_2__=_01_c_drive\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER #NFS_
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
				ELEMENT7=_%NET_PATH%.TXT

				Array_FileName%ArrayCount% =%ELEMENT1%%ELEMENT2%%ELEMENT3%%ELEMENT4%%ELEMENT5%%ELEMENT7%
				MSGBOX % Array_FileName%A_Index%
			}
		}

		Loop %ArrayCount%
		{
			file := FileOpen(Array_FileName%A_Index%, "w")
			if IsObject(file)
			{
				MSGBOX % Array_FileName%A_Index%

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


