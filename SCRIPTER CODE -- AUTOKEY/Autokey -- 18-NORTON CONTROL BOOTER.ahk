;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 18-NORTON CONTROL BOOTER.ahk
;# __ 
;# __ Autokey -- 18-NORTON CONTROL BOOTER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# 1ST START TIME [ Wed 02:32:20 Am_15 Nov 2017 ]
;# 2ND END   TIME [ Wed 04:14:50 Am_15 Nov 2017 ]
;# 3RD FINAL TIME [ Wed 07:00:40 Am_15 Nov 2017 ]
;# 4TH FROM  TIME [ Sat 16:00:22 Pm_21 Apr 2018 ]
;# 4TH TO    TIME [ Sat 19:10:00 Pm_21 Apr 2018 ] 3 HOUR _ Skillful Improver
;# 5TH FROM  TIME [ Sun 23:00:00 Pm_22 Apr 2018 ] 32 Bit 64 Bit Sort Out
;# 5TH TO    TIME [ Sun 23:59:00 Pm_22 Apr 2018 ] 
;# 6TH FROM  TIME [ Mon 15:40:00 Pm_23 Apr 2018 ] Shortcut Desktop Norton 
;# 6TH TO    TIME [ Mon 15:59:00 Pm_23 Apr 2018 ] 
;# 7TH FROM  TIME [ Mon 18:00:00 Pm_23 Apr 2018 ] Minimize Amount Code Searcher Neater
;# 7TH TO    TIME [ Mon 18:40:00 Pm_23 Apr 2018 ] 
;# 8TH FROM  TIME [ Mon 23:36:14 Pm_30 Apr 2018 ] Better Code Detect Colour Green for Done
;# 8TH TO    TIME [ Tue 06:50:00 Pm_30 Apr 2018 ] 7 Hour 20 Minute _ 31 URL's Open
;# 
;====================================================================

;# ------------------------------------------------------------------
; An AutoHotKeys Project to Start Norton for a Quick Scan 
; Every Time Boot Up
; It Got Folder Scanning to Find the Up to Date Norton to Use
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/bsvrrmyknxr7pvd/Autokey%20--%2018-NORTON%20CONTROL%20BOOTER.ahk?dl=0
;# ------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
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

DetectHiddenWindows, on
;DetectHiddenWindows, off

SetStoreCapslockMode, off

PASS_LEVEL_VAR_1=FALSE
PASS_LEVEL_VAR_2=FALSE
O_PASS_LEVEL_VAR_2=FALSE
O_COLOR=0
O_FlipFlop_MinMax=-2
Counter_MinMax=0
Counter_Minimze_2=14
SET_GO_ACTIVATE=1

global KEEP_COUNTER_DOWN_Counter_Minimze_2=0
global HWND_1=0
global HWND_2=0
global COLOR=0
global O_COLOR

EXIT_COUNTER=0

Counter_MinMax=0

CoordMode Mouse, SCREEN
CoordMode Pixel, SCREEN

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

setTimer TIMER_PREVIOUS_INSTANCE,1

GOSUB RUN_NORTON

setTimer TIMER_SUB_1,1000
setTimer TIMER_COUNT_EXIT,1000
setTimer SPEED_CHECK_NORTON_COLOR,1
setTimer SPEED_GET_NORTON_COLOR,1
SetTimer TIMER_TEST,off

RETURN

TIMER_TEST:    
    WinGet, HWND, ID, A  ; Get Active Window
    WinGet, OutputVar, ControlList, ahk_id %HWND%
    Tooltip, % OutputVar ; List All Controls of Active Window
return

RUN_NORTON:

; ----------------------------------------
; SLEEP FOR A BIT IF JUST BOOTER
; NORTON NEEDS IT TO GET ENGINE GOING
; BEFORE LAUNCHER OR FAIL WITH NOT RESPOND
; ----------------------------------------
; IN THE FIRST FIVE MINUTES OF TICKER
; ----------------------------------------
COUNT_TICK_TIME=% 1000*60*4
LOOP
{
IF A_TICKCOUNT< %COUNT_TICK_TIME%
	SLEEP 1000
IF A_TICKCOUNT> %COUNT_TICK_TIME%
	BREAK
;TOOLTIP , %A_TICKCOUNT% __ %COUNT_TICK_TIME%
}
	
SET_GO=TRUE
FILENAME="" 

;DetectHiddenWindows, on
;WinWait, ahk_class NortonMainWindow, , 120
;loop 40
;{
;	sleep 1000 
;	if WinExist("ahk_class NortonMainWindow")
;		break
;}
;msgbox yes
;SoundBeep , 2000 , 100
;SoundBeep , 2500 , 100

;nortonsecurity.exe

IfWinExist ahk_class SymHTMLDialog 
	SET_GO=FALSE
IfWinExist ahk_class Sym_Common_Scan_Window
	SET_GO=FALSE

	
	
IF SET_GO=TRUE
{
	ProgramFilesX86 := A_ProgramFiles . (A_PtrSize=8 ? " (x86)" : "")

	; 32 bit Norton Installed
	IF FILENAME="" 
	{	
		Loop Files, %ProgramFilesX86%\Norton Security with Backup\Navw32.exe, R  ; Recurse into subfolders.
		{
			FILENAME = %A_LoopFileFullPath%
			SHORTCUT_FILENAME = %A_LoopFileDir%\uistub.exe
		}
	}
	; 64 bit Norton Installed
	IF FILENAME="" 
	{
		Loop Files, %A_ProgramFiles%\Norton Security with Backup\Navw32.exe, R  ; Recurse into subfolders.
		{
			FILENAME = %A_LoopFileFullPath%
			SHORTCUT_FILENAME = %A_LoopFileDir%\uistub.exe
		}
	}
	
	; -------------------------------------------------------------
	; DOUBLE BACKUP IF NORTON IS NOT FOUND
	; THEY MAY OF CHANGED THE FOLDER NAME KEPT AT
	; MAKE SOME MORE SOUND TO INDICATE SOMETHING IS UP 
	; FIND THE CORRECT NEAREST LOCATION OF NORTON AND EDIT IN ABOVE
	; -------------------------------------------------------------
	
	; 32 bit Norton Bigger Search
	IF FILENAME="" 
	{
		SoundBeep , 2000 , 100
		SoundBeep , 2500 , 100
		SoundBeep , 3500 , 100

		Loop Files, %ProgramFilesX86%\Navw32.exe, R  ; Recurse into subfolders.
		{
			FILENAME = %A_LoopFileFullPath%
			SHORTCUT_FILENAME = %A_LoopFileDir%\uistub.exe
		}
	}
	; 64 bit Norton Bigger Search
	IF FILENAME="" 
	{
		SoundBeep , 2000 , 100
		SoundBeep , 2500 , 100
		SoundBeep , 3500 , 100

		Loop Files, %A_ProgramFiles%\Navw32.exe, R  ; Recurse into subfolders.
		{
			FILENAME = %A_LoopFileFullPath%
			SHORTCUT_FILENAME = %A_LoopFileDir%\uistub.exe
		}
	}

	IF FILENAME="" 
	{
		SoundBeep , 3500 , 100
		SoundBeep , 2500 , 100
		MSGBOX, "AutoHotkeys __ Norton Security Was Not Found - ENDER"
		SoundBeep , 3500 , 100
		SoundBeep , 2500 , 100
		Exitapp
	}
	
	IF FILENAME<>"" 
	{
		SoundBeep , 2000 , 100
		
		;if FileExist("%A_Desktop%\Norton Security.lnk")=false
		
		FileCreateShortcut, %SHORTCUT_FILENAME%, %A_Desktop%\Norton Security.lnk, , "%A_ScriptFullPath%"

		;-------------------------------------------------------
		; NORTON TAKES A DELAY AT START UP TO GET RUNNING PROPER
		; FAILS TO LAUNCH TOO EARLY
		;-------------------------------------------------------
		Set_Go=False
		loop, 120
		{
			sleep 1000
			Process, Exist, %FILENAME%
			If Not ErrorLevel
			{
				Run, "%FILENAME%" /UPDATE /QUICK , , , OutputVarPID
				;msgbox %OutputVarPID%
				if OutputVarPID>0
				{		
					Set_Go=True
					break
				}
				ELSE
				{
					MSGBOX, NORTON RETURNING PID = 0
				}
			}
		}
		if Set_Go=False
		{
			msgbox Find HWND Timed out. Waiting for Norton to Loadup. End Script.
			Exitapp
		}
		
		;WinWait, ahk_class SymHTMLDialog, , 1200
		;if ErrorLevel
		;{
		;	MsgBox, WinWait timed out. Waiting for Norton to Loadup. End Script.
		;	Exitapp
		;}
	}
}
RETURN

TIMER_SUB_1:

DetectHiddenWindows, oFF

	WinGet, HWND_1, ID, ahk_class SymHTMLDialog
	IF (HWND_1>0 and PASS_LEVEL_VAR_2="FALSE")
	{

		; ABOVE WIN XP
		WINCLOSE ahk_id %HWND_1%

		if WinExist("ahk_id " HWND_1)>0
		{
			; WIN XP MODE
			WinActivate, ahk_class SymHTMLDialog
			Sendinput !{F4}
		}
		SoundBeep , 2500 , 100
		WinWait, Quick Scan ahk_class Sym_Common_Scan_Window, , 80
		WINMINIMIZE
		; UNLESS USE WINWAIT HERE DON'T GET PROPER
		; MINIMIZE HAPPEN LATER
		SoundBeep , 2500 , 100
		
		if WinExist("ahk_id " HWND_1)=0
			PASS_LEVEL_VAR_1=TRUE
		
	}
		
	WinGet, HWND_2, ID, Quick Scan ahk_class Sym_Common_Scan_Window

	IF HWND_2>0 
		{
		PASS_LEVEL_VAR_1=TRUE
		PASS_LEVEL_VAR_2=TRUE
		}
	
	;OFF TO A QUICK START
	IF O_PASS_LEVEL_VAR_2<>%PASS_LEVEL_VAR_2%
	{
		WinGet, Current_MinMax, MinMax, ahk_id %HWND_2%
		GOSUB HAS_NORTON_FINISHED
	}
	O_PASS_LEVEL_VAR_2=%PASS_LEVEL_VAR_2%
	
	WinGetClass, This_Class, ahk_id %HWND_2%
	WinGet, Current_MinMax, MinMax, ahk_id %HWND_2%

	;and COLOR=0x23A31C) ; FINISHED COLOR -- GREEN
	;and COLOR=0xFDC649) ; WORKING  COLOR -- ORANGE YELLOW
	;and COLOR=0xFC6120) ; ABORT    COLOR -- RED

	FlipFlop_MinMax=%Current_MinMax%
	if O_FlipFlop_MinMax=-2
		O_FlipFlop_MinMax=%Current_MinMax%
	
	if FlipFlop_MinMax<>%O_FlipFlop_MinMax%
	if FlipFlop_MinMax=0
	{
		Counter_MinMax+=1
		Counter_Minimze_2=14
		SET_GO_ACTIVATE=1
	}
	
	IF O_COLOR<>%COLOR%
		Counter_Minimze_2=14
		
	O_COLOR=%COLOR%
	
	IF (COLOR=0x23A31C)
		Counter_MinMax=1
	
	O_FlipFlop_MinMax=%Current_MinMax%

	If (Current_MinMax=0
	and Counter_MinMax<2
	and Counter_Minimze_2<1 
	and Counter_Minimze_2>-5)
	{
		GOSUB HAS_NORTON_FINISHED
		SoundBeep , 2500 , 100
		SoundBeep , 3500 , 100
	}

	If Counter_MinMax>0 
	{
		Counter_Minimze_2-=1
		if Counter_Minimze_2<-100
			Counter_Minimze_2=-100
	}
	
	If (PASS_LEVEL_VAR_1="TRUE" 
	and PASS_LEVEL_VAR_2="TRUE" 
	and (!HWND_1) 
	and (!HWND_2))
		{
		SoundBeep , 3500 , 100
		SoundBeep , 2500 , 100
		Exitapp
		}
	
	If Counter_Minimze_2<-20
		{
		GOSUB HAS_NORTON_FINISHED
		IF COLOR=0x23A31C ; GREEN
		{
			SoundBeep , 3500 , 100
			SoundBeep , 2500 , 100
			Exitapp
		}
		}

	GOSUB SPEED_CHECK_NORTON_COLOR

	if (!COLOR)
		GOSUB SPEED_GET_NORTON_COLOR

	IF (Current_MinMax=0 
	and HWND_2>0
	and SET_GO_ACTIVATE=1)
	{
		;IF (COLOR<>0x23A31C)
			;WinActivate, ahk_class Sym_Common_Scan_Window
		;IF (!COLOR)
			;WinActivate, ahk_class Sym_Common_Scan_Window
		SET_GO_ACTIVATE=0
	}
		
	If Counter_Minimze_2<-20
	and Current_MinMax=0
	and SET_GO_ACTIVATE=1)
	{
		;IF (COLOR<>0x23A31C)
			;WinActivate, ahk_class Sym_Common_Scan_Window
		;IF (!COLOR)
			;WinActivate, ahk_class Sym_Common_Scan_Window
		SET_GO_ACTIVATE=0
	}
RETURN

TIMER_COUNT_EXIT:
	EXIT_COUNTER+=1
	; 
	; 3600 1 HOUR
	; 7200 2 HOUR
	;
	IF EXIT_COUNTER>3600 ; 1 HOUR -- SOMETIMES STUCK NEVER EXIT HERE -- SPEED_GET_NORTON_COLOR:
	{
		EXITAPP
	}
RETURN

SPEED_GET_NORTON_COLOR:

	IF (!HWND_2)
		RETURN
    
	WinGet, HWND_2, ID, Quick Scan ahk_class Sym_Common_Scan_Window
	IF HWND_2>0
	{
		
		WinGetPos , X, Y, Width, Height, ahk_id %HWND_2%
		;MSGBOX, %X% __ %Y%
		X+=4
		Y+=88 ; 81 SET FIRST FIND VALUE ADD EXTRA FOR MOVEMENT PLENTY OF ROOM
		IF COLOR<>"0x23A31C"
			{
			PixelGetColor Color, %X%, %Y%, RGB
			}
	}
RETURN

SPEED_CHECK_NORTON_COLOR:
	IF (!HWND_2)
		RETURN

	; IF SCAN BEEN ABORTED OR MAYBE VIRUS FOUND QUIT
	WinGet, HWND_2, ID, Quick Scan ahk_class Sym_Common_Scan_Window
	IF HWND_2>0
	{
		WinGetPos , X, Y, Width, Height, ahk_id %HWND_2%
		X+=4
		Y+=88 ; 81 SET FIRST FIND VALUE ADD EXTRA FOR MOVEMENT PLENTY OF ROOM
		IF COLOR<>"0x23A31C"
			{
			PixelGetColor Color, %X%, %Y%, RGB
			}

		IF COLOR=0xFC6120
		{
			SoundBeep , 3500 , 100
			SoundBeep , 2500 , 100
			Exitapp
		}
		;FC6120 (Red=FC Green=61 Blue=20)
		;MOSTLY RED
	}
RETURN

HAS_NORTON_FINISHED:
; HAS IT FINISHED BY SHOW GREEN COLOUR		
; ONLY DETECTABLE WHEN WINDOW NOT MINIMIZED
; WHEN FINISHED WINDOW POP'S UP AGAIN
; THE ONE LAST TIME FOR THIS CODE
; SETTING MINIMIZE HAPPENS TWICE FIRST 
; STRAIGHT AWAY THEN THEN SECOND 
; AFTER TEN SECONDS
; THAT IS TO SHOW IF THE USER WANTED IT TO POP UP 
; AGAIN AFTER FIRST TIME AND DOES SO MANUALLY
; THERE IS STILL ONE IN QUE TO DO AGAIN
; MAYBE SHOULD ABORT THAT LOT AND JUST DO ONE
; BUT YOU MAY WANT A LOOK AT IT AND DECIDE TO WALK OFF 
; AND NOT WANT IT SHOWING THERE WHEN COME AGAIN
; IF YOU COME BACK TO IT AFTER A LONG TIME AND POP THE WINDOW UP MANUALLY
; YOU DON'T WANT IT SUDDENLY POPPED DOWN AGAIN
; END OF DAY WINDOW POPS UP AGAIN BUT IT DON'T LEARN IF GREEN AGAIN 
; UNLESS ACTIVE
; WHICH MIGHT BE WHAT WE WANT IF ANOTHER WINDOW NOT GOT IN THE WAY
; MANUAL MIN AND MAX BUMP UP THE COUNTDOWN TIMER TO +14
; AT THE END IF PROGRAM STILL RUN IT WILL MINIMIZE THE 
; FINISH GREEN RESULT BY NORTON
; NOT ANOTHER WAY AROUND TOO MUCH DETECTION ON MANUAL INTERVENTION 
; WHILE GREEN AT END IS NOT WHAT WANTER
; ------------------------------------
IF HWND_2>0 
{
	;IF SCAN BEEN ABORTED OR MAYBE VIRUS FOUND QUIT
	WinGet, HWND_2, ID, Quick Scan ahk_class Sym_Common_Scan_Window
	IF HWND_2>0
	{
	WinGetPos , X, Y, Width, Height, ahk_id %HWND_2%
	X+=4
	Y+=88 ; 81 SET FIRST FIND VALUE ADD EXTRA FOR MOVEMENT PLENTY OF ROOM
	IF COLOR<>0x23A31C
	{
		PixelGetColor Color, %X%, %Y%, RGB
	}

	IF COLOR=0x23A31C
	{
		WinMinimize, ahk_class Sym_Common_Scan_Window	
		;WinMinimize  ahk_id %HWND_2%
		SoundBeep , 3500 , 100
		SoundBeep , 2500 , 100
		Exitapp
	}
	ELSE
	{
		IF COLOR<>%O_COLOR%
		{
			WinMinimize, ahk_class Sym_Common_Scan_Window	
			; SOUND KEEPS BEEPING AT END IN GREEN DONE _ SORT WITH O_COLOR
			SoundBeep , 3500 , 100
			;WinMinimize  ahk_id %HWND_2%
		}
	}		
	O_COLOR=%COLOR%
	}
}
Return


IS_WINDOW_MINIMIZED(HWND)
{
LOOP {
	WINHIDE ahk_id %HWND%
	WinGet, Style, Style, ahk_id %HWND%
	; 0x10000000 is WS_MINIMIZED  Style & 0x10000000 = 0 IS HIDDEN > 0 NOT HIDDEN
	IF (Style & 0x10000000)>0
	{
	SLEEP 10
	WINHIDE ahk_id %HWND%
	}
} UNTIL (Style & 0x10000000)=0
}
RETURN

MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
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
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
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
;# -----------------------------------------------------------------0
; exit the app


; GOOD SCRIPT EXAMPLE PAGE HELPER
; ----
; 1 Hour Software by Skrommel - DonationCoder.com
; http://www.donationcoder.com/Software/Skrommel/index.html#GoneIn60s
; ----

; Junk Code Been Try
;Is64Bit() 
;{
   ;return (RegexMatch(EnvGet("Processor_Identifier"), "^[^ ]+64") > 0)
;}


;--------------------------------------------------------------------
; FUNCTION SET 
;--------------------------------------------------------------------

IS_WINDOW_HIDDEN_AND_HIDE(HWND)
{
LOOP {
	WINHIDE ahk_id %HWND%
	WinGet, Style, Style, ahk_id %HWND%
	; 0x10000000 is WS_VISIBLE  Style & 0x10000000 = 0 IS HIDDEN > 0 NOT HIDDEN
	IF (Style & 0x10000000)>0
	{
	SLEEP 10
	WINHIDE ahk_id %HWND%
	}
} UNTIL (Style & 0x10000000)=0
}
RETURN

;--------------------------------------------------------------------
;--------------------------------------------------------------------

