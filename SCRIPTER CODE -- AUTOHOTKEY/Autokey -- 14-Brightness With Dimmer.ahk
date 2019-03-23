;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AutoKey -- 14-Class_Monitor-master\src\AutoKey -- 14-Brightness.ahk
;# __ 
;# __ AutoKey -- 14-Brightness.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# Fri 15 September 2017 & Saturday
;# __ 
;  =============================================================

; GLOBAL SETTINGS ===================================================

#Warn
#NoEnv
#SingleInstance Force
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


#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer\Class_Monitor_Master\SRC\Class_Monitor.ahk     ; include the class here

; ------------------------------------------------------------------
; ----
; GitHub - jNizM/Class_Monitor: Monitor Class (WinAPI)
; https://github.com/jNizM/Class_Monitor
; ----
; ----
; Class Monitor (Brightness, ColorTemperature) - AutoHotkey Community
; https://autohotkey.com/boards/viewtopic.php?f=6&t=7854
; ----
; ------------------------------------------------------------------

; HITT CLONE OR DOWNLOAD __ DON'T HAVE TO BE A SIGNED UP MEMBER
; Class_Monitor.ahk __ PUT THIS IN A FOLDER WITH THIS CODE OR VICE VERSA YOUR CODE HERE IN THIS FOLDER
; UNLIKE THE DIM CONTROL OF COMPUTER STANDBY THAT NOT WORKING WELL ABOVE WINDOWS 7
; YOU REQUIRE SECPOL.MSI TO RUN WHICH REQUIRE PRO VERSION NOT HOME VERSION
; ALSO HERE __ THIS SET THE BRIGHTNESS OF ALL MONITOR
; BUT EXTERNAL MONITOR CAN EASY BE SWITCHED OFF AT NIGHT OTHER LEFT ON AND BRIGHT ON YOUR EYE
; THAT IS GOOD FOR YOU
; ALSO UNLIKE DIM CONTROL ALL MONITOR IS ANSWER IT SET A DIFFERENT KIND OF BRIGHTNESS 
; ADJUSTMENT _ I THINKER __ THAT WHY ALL MONITOR GET SET INCLUDE HDMI ON 
; WHERE DIM CONTROL USUAL NOT AFFECTING
; BUT HERE CONTROL SCAN ALL MONITOR
; SEEM BUG IN CODE AS NOT ABLE USE PART WHERE A CERTAIN ONE MONITOR
; GOOD NIGHT OVER AND OUT 
; Fri 15 September 2017 01:09:04----------
; 
; TO LATER ADD WORK TO DO _ NOT DIM SCREEN SAVER WHEN CERTAIN ACTIVITY PROGRAM RUNNER
; DONE
; TO LATER ADD WORK TO DO _ SEND A DIMMER WHEN MOUSE CLICK OVER ONE SIDE OF THE SCREEN 
; FOR A TIME OR OFF AGAIN
;#
; SOME OF THE IDLE ROUTINE ABOUT WERE CODED IN 2005 WHEN WINDOWS_98 WAS AROUND
; THESE DIMMER ROUTINE ARE CURRENTLY
;#
; ENERGY PROBLEM OVER
; ------------------------------------------------------------------
; WHY PAY FOR PRO TO HAVE ENERGY FEATURE
; --------------------------------------
; 1ST EVENT OF DIMMER IS IN 4 SECOND IDLE NEXT AFTER ARE IN 2 MINUTE
; ------------------------------------------------------------------
; WORK HISTORY
; ------------------------------------------------------------------

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; Fri 15 September 2017 01:09:04---- ABOVE WORK BEGINNER
; ------------------------------------------------------------------

; -------------------------------------------------------------------
; 002 ---------------------------------------------------------------
; Wed 21 March     2018 02:10:21---- Work to Detect Explorer Was Active ; As it Remains Present When Closed and Minimal Sometimes Error There 
; had to be Resolved _ Explorer and Other Program make Quicker Timeout ; Brightness Change _ Few Hour-ing
; ------------------------------------------------------------------

; -------------------------------------------------------------------
; 003 ---------------------------------------------------------------
; WORK EXTRA INCLUDE DAY TIME MODE BRIGHTNESS UP
; -------------------------------------------------------------------
; FROM TIME __ Wed 20-Jun-2018 15:20:52
; TO   TIME __ Wed 20-Jun-2018 15:54:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 004 ---------------------------------------------------------------
; WORK _ INCLUDE A BLACK SCREEN ALSO WITH DIM _ SLEEP BETTER AT NIGHT IN ROOM
; -------------------------------------------------------------------
; FROM TIME __ Tue 03-Jul-2018 22:01:03
; TO   TIME __ Wed 04-Jul-2018 01:24:25 4 HOUR 24 MIN FOR A SMALL 
;                                       PIECE OF CODE BOO 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 005 ---------------------------------------------------------------
; WORK _ ADD DISPLAY OFF WORKING POWER ENERGY SAVER COMBINED WITH BLANK SCREEN
; -------------------------------------------------------------------
; FROM TIME __ Wed 04-Jul-2018 07:46:36 __ SESSION 01 OF 02
; TO   TIME __ Wed 04-Jul-2018 08:30:44

; FROM TIME __ Wed 04-Jul-2018 17:42:15 __ SESSION 02 OF 02
; TO   TIME __ Wed 04-Jul-2018 18:19:19
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 00* ---------------------------------------------------------------
; MORE WORK _ DEBUG NOT WORK CORRECTLY TIMER AND LATCH 
; -------------------------------------------------------------------
; FROM TIME __ Wed 04-Jul-2018 22:35:01 _ DEBUG VARIABLE SYNTAX
; TO   TIME __ Wed 04-Jul-2018 23:22:00 _ 50 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 007 ---------------------------------------------------------------
; MORE WORK _ THE LAST OF THE LATCHING BUGGS SORTED
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 10:06:11 _ DEBUG VARIABLE SYNTAX
; TO   TIME __ Thu 05-Jul-2018 10:24:00 _ 18 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 008 ---------------------------------------------------------------
; MORE WORK _ MY 7-ASUS HAS A DUPLICATE MONITOR SYSTEM AND MONITOR 
; POWER LEAVES DUPLICATE MODE SO ONLY BLANK SCREEN USER
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 22:22:00
; TO   TIME __ Thu 05-Jul-2018 22:25:00 _ 3 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 009 ---------------------------------------------------------------
; MORE WORK _ FINIAL LOOK AT CODE PROBLEM MONITOR WASN'T COMING OUT OF 
; STANDBY AT THE PRESET TIME IN MORNING _ ANSWER WAS TO MOVE A THE MOUSE
; A LITTLE BIT-PIXEL AND BACK AGAIN REMARK COMMENTS ARE NEAR CODE
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 22:20:55
; TO   TIME __ Thu 05-Jul-2018 12:02:00 _ 1 HOUR 42 MINUTE
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location Internet
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
; https://www.dropbox.com/s/qezyk9thq0boir2/Autokey%20--%2014-Brightness%20With%20Dimmer.ahk?dl=0
; ------------------------------------------------------------------


;#InstallKeybdHook
;#InstallMouseHook

; IF YOU GOT ANYTHING KEYBOARD HOOKING SENDING A KEY OR MOUSE
; AND USE HERE A_TimeIdlePhysical __ INSTEAD OF A_TimeIdle
; ONLY Physical INPUT NOT SIMULATED SENDER

; TRICK PART OF CODE IS MANY DON'T TELL HAVE TO FIND IDLE AS WELL ALSO THE ACTIVE PART
; ALL VERY WELL SWITCH THE SCREEN SAVER ON BUT____ AND THEN OFF SCREEN SAVER


OnExit, EOF                         ; set onexit to the label "EOF"
Display := New Monitor()            ; initialize / start the class

; SCRIPT ============================================================

DetectHiddenWindows, on

CoordMode, Mouse, Screen
Mouse_Idle = 0 
Mouse_Idle_Flip_Flop_Toggle := "True"
LastX = 0
LastY = 0
VAR_A__TimeIdle = 0
VAR_Z__TimeIdle_1 = 4000
VAR_Z__TimeIdle_4_DEFAULT = 80000 ; 4 MINUTE
VAR_Z__TimeIdle_3_FORCE = 2000 
VAR_Z__TimeIdle_2 = %VAR_Z__TimeIdle_2_DEFAULT%
VAR_Z__TimeIdle = %VAR_Z__TimeIdle_1%
OLDWinActive = 0
WinActive_2 = 0

GLOBAL IN_DAY
GLOBAL O_IN_DAY_1
GLOBAL O_IN_DAY_2
GLOBAL BLANK_DIMMER
GLOBAL BLANK_DIMMER_VAR
GLOBAL BLANK_DIMMER_TIME
GLOBAL ALLOW_DIMMER

ALLOW_DIMMER := "True"

O_IN_DAY_1=FALSE
BLANK_DIMMER_VAR=FALSE

BLANK_DIMMER_TIME=80
BLANK_DIMMER:= A_Now
BLANK_DIMMER+= %BLANK_DIMMER_TIME%, Seconds


; IF (A_ComputerName="7-ASUS-GL522VW")
	; PAUSE


SoundBeep , 1000 , 100
SoundBeep , 3000 , 100


; IF (A_ComputerName="1-ASUS-X5DIJ")
	; PAUSE
; IF (A_ComputerName="2-ASUS-EEE")
	; PAUSE
; IF (A_ComputerName="4-ASUS-GL522VW")
	; PAUSE
	
Gui, Color, black
Gui +AlwaysOnTop
Gui -Caption
; Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%	
Gui, hide
; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
SendMessage, 0x112, 0xF170, 0,, Program Manager


GOSUB, MONITOR_BRIGHTNESS_UP


SetTimer,Mouse_Idle_Timer, 1000     ; Check Every Second
setTimer TIMER_PREVIOUS_INSTANCE,1


; ------------------------------------
Mouse_Idle_Timer:


COUNT_TICK_TIME=% 1000*60*24
; TOOLTIP %A_TICKCOUNT% __ %COUNT_TICK_TIME%

IF A_TICKCOUNT< %COUNT_TICK_TIME%
	RETURN

GOSUB, MONITOR_BRIGHTNESS_DIMMER_PER_DAY

DetectHiddenWindows, on
ALLOW_DIMMER := "True"

if WinActive("ahk_class MediaPlayerClassicW")
ALLOW_DIMMER := "False"

; if WinActive("ahk_class Notepad++")
; ALLOW_DIMMER := "False"

If WinActive("tube - Google Chrome")
	ALLOW_DIMMER := "False"
	
If WinActive("YouTube - Google Chrome")
	ALLOW_DIMMER := "False"

IF (A_ComputerName="1-ASUS-X5DIJ")
	ALLOW_DIMMER := "False"
IF (A_ComputerName="2-ASUS-EEE")
	ALLOW_DIMMER := "False"
IF (A_ComputerName="3-LINDA-PC")
	ALLOW_DIMMER := "False"
IF (A_ComputerName="5-ASUS-P2520LA") 
	ALLOW_DIMMER := "False"
;IF (A_ComputerName="4-ASUS-GL522VW")
	;ALLOW_DIMMER := "False"
IF (A_ComputerName="8-MSI-GP62M-7RD")
	ALLOW_DIMMER := "False"
;IF (A_ComputerName="7-ASUS-GL522VW")
;	ALLOW_DIMMER := "False"

GOSUB IS_IN_DAY


isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
if isFullScreen 
IF TRUE = FALSE
{

	HWND := WinExist("A")
	; WinGet, HWND, ID, "A"
	WinGetClass, This_Class_2, ahk_id %HWND%
	WinGet, Path_2, ProcessName, ahk_id %HWND%
	WinGetTitle, OutputVar_1, ,ahk_id %HWND%
	WinGetText, OutputVar_2, ahk_id %HWND%
	
	SET_GO=0
	IF Path_2=explorer.exe
		SET_GO=0
	IF This_Class_2=Progman
		SET_GO=1
	IF Path_2=notepad++.exe
		SET_GO=1
	IF Path_2=TeamViewer.exe
		SET_GO=1
	IF Path_2=i_view32.exe
		SET_GO=1
	IF Path_2=chrome.exe
		SET_GO=1
	IF Path_2=mpc-hc64.exe
		SET_GO=1
	IF (A_ComputerName<>"7-ASUS-GL522VW")
		SET_GO=1

	if SET_GO=0
		{
			SOUNDBEEP 1000,200
			MSGBOX , CLASS__ %This_Class_2% `r`nPROCESS__ %Path_2% `r`nTITLE__ %OutputVar_1%`r`nTEXT__ %OutputVar_2%
		}
}


isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
if isFullScreen 
	{
		ALLOW_DIMMER := "False"
		;MSGBOX % ALLOW_DIMMER
	}
	ELSE
	{
		;SoundBeep , 2000 , 100
	}
 
SetTitleMatchMode, 2
FORCE_DIMMER_QUICKER := "FALSE"
VAR_Z__TimeIdle_2=%VAR_Z__TimeIdle_4_DEFAULT%
DetectHiddenWindows, off
WinActive_2 = 0

SET_GO=FALSE
IF (A_ComputerName="7-ASUS-GL522VW")
	SET_GO=TRUE

IF SET_GO=TRUE
{
	if WinActive("GoodSync")
	{
		VAR_Z__TimeIdle_2=%VAR_Z__TimeIdle_3_FORCE%
		VAR_Z__TimeIdle = %VAR_Z__TimeIdle_2%
		WinActive_2 = 1
	}

	if WinActive("Notepad++")
	{
		VAR_Z__TimeIdle_2=%VAR_Z__TimeIdle_3_FORCE%
		VAR_Z__TimeIdle = %VAR_Z__TimeIdle_2%
		WinActive_2 = 1
	}
	if WinActive("ahk_class Chrome_WidgetWin_1")
	{
		VAR_Z__TimeIdle_2=%VAR_Z__TimeIdle_3_FORCE%
		VAR_Z__TimeIdle = %VAR_Z__TimeIdle_2%
		WinActive_2 = 1
	}

	; EXPLORER	
	; --------
	#IfWinActive ahk_class CabinetWClass
	{
		isWindowFocusExplorer_2 := isWindowFocusExplorer( "A" ) ; ActiveWindow
		if isWindowFocusExplorer_2
		{
			VAR_Z__TimeIdle_2=%VAR_Z__TimeIdle_3_FORCE%
			VAR_Z__TimeIdle = %VAR_Z__TimeIdle_2%
			WinActive_2 = 1
		}
	}
}

IF (OLDWinActive=0 and WinActive_2=1)
	{ 
		;SoundBeep , 2000 , 100
		GOSUB, MONITOR_BRIGHTNESS_DIM
	} 
		
OLDWinActive=%WinActive_2%
 
GetKeyState, state, LButton
if state = D              
	ALLOW_DIMMER := "False"
	; MOUSE BUTTON LEFT HELD DOWN WHEN DRAGGER FOR LONG NOT DETECT BY IDLE ACTIVE UNLESS SWITCH
 
 
;#-------------------------------
If (A_TimeIdle > VAR_Z__TimeIdle and ALLOW_DIMMER = "True")
{
	VAR_Z__TimeIdle = %VAR_Z__TimeIdle_2%
	GOSUB, MONITOR_BRIGHTNESS_DIM
}
GOSUB, Keyboard_Idle_Timer
;#-------------------------------
Return ; End of Mouse_Idle_Timer

; ------------------------------------------------------------------
Keyboard_Idle_Timer:
; Tooltip % + A_TimeIdle " -- " VAR_A__TimeIdle 
;TEST DEBUG ___________

IF (A_TimeIdle < VAR_A__TimeIdle)
{
	;SoundBeep , 2500 , 100
	;TEST DEBUG ___________
	GOSUB, MONITOR_BRIGHTNESS_UP

	BLANK_DIMMER:= A_Now
	BLANK_DIMMER+= %BLANK_DIMMER_TIME%, Seconds

}
VAR_A__TimeIdle = %A_TimeIdle%

; A_TimeIdle - SHOW TIME SINCE LAST KEYBOARD OR MOUSE IN MILLISECOND
; THE DETECT IS IF LOWER THAN
; ---------------------------


return

; ------------------------------------------------------------------
MONITOR_BRIGHTNESS_DIM:

IF ALLOW_DIMMER=FALSE
	RETURN
	
If (Mouse_Idle_Flip_Flop_Toggle = "False")
{
	; BRIGHTNESS LOW 1 IS MINIMAL DIMMER
	IF WinActive_2=0
	{ 
		SoundBeep , 1500 , 100
	}
	Monitor.SetBrightness(1, 1, 1)
	SetTimer,Mouse_Idle_Timer, 10 ; RAPID RESPONSE BRING OUT OF DIM
	; SET TIMER TO THE SECOND 1000MS FOR EASIER COUNTER TIME_OUT
	Mouse_Idle_Flip_Flop_Toggle := "True"
}
RETURN

; ------------------------------------------------------------------
MONITOR_BRIGHTNESS_UP:
If (Mouse_Idle_Flip_Flop_Toggle = "True")
{
	;SoundBeep , 2500 , 100
	; BRIGHTNESS UP 50% IS 100% __ Full Bright
	Monitor.SetBrightness(127, 127, 127)
	SetTimer,Mouse_Idle_Timer, 1000
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
	SendMessage, 0x112, 0xF170, 0,, Program Manager
	Gui, HIDE
	Mouse_Idle_Flip_Flop_Toggle := "False"
}
RETURN

MONITOR_BRIGHTNESS_DIMMER_PER_DAY:
	
	SET_GO=FALSE
	IF (A_ComputerName="1-ASUS-X5DIJ")
		SET_GO=TRUE
	IF (A_ComputerName="2-ASUS-EEE")
		SET_GO=TRUE
	IF (A_ComputerName="3-LINDA-PC")
		SET_GO=TRUE
	IF (A_ComputerName="4-ASUS-GL522VW")
		SET_GO=TRUE
	IF (A_ComputerName="5-ASUS-P2520LA") 
		SET_GO=TRUE
	IF (A_ComputerName="7-ASUS-GL522VW")
		SET_GO=TRUE
	IF (A_ComputerName="8-MSI-GP62M-7RD")
		SET_GO=TRUE

	IF SET_GO=FALSE
		RETURN

	IF A_NOW<%BLANK_DIMMER%
	{
		BLANK_DIMMER_VAR=TRUE
		RETURN
	}
	
	SET_GO=FALSE
	IF BLANK_DIMMER_VAR=TRUE
	{
		SET_GO=TRUE
		; AFTER MOUSE ACTIVITY LET THE CODE THINK NEXT IS WANT HIDE MONITOR
		BLANK_DIMMER_VAR=FALSE
	}
	
	GOSUB IS_IN_DAY
	
	IF IN_DAY=TRUE
		SET_GO=FALSE
	
	; TOOLTIP % IN_DAY " __ " 
	
	
	ALLOW_DIMMER := "True"
	isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
	if isFullScreen 
	{
		ALLOW_DIMMER := "False"
	}
	
	IF (O_IN_DAY_1<>%IN_DAY% or SET_GO="TRUE")
	{
	
		POWER_SCREEN_SAVE_OFF := "False"
		IF (A_ComputerName="7-ASUS-GL522VW")
			POWER_SCREEN_SAVE_OFF := "True"
		IF (A_ComputerName="8-MSI-GP62M-7RD")
			POWER_SCREEN_SAVE_OFF := "True"
		
		DIMMER_ONLY_NOT_BLANK := "False"
		IF (A_ComputerName="8-MSI-GP62M-7RD")
			DIMMER_ONLY_NOT_BLANK := "True"
	
		IF (IN_DAY="FALSE" or SET_GO="TRUE")
			IF (ALLOW_DIMMER="True")
			{
				IF DIMMER_ONLY_NOT_BLANK="False"
				{
					Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%	
				}
				ELSE
				{
					Monitor.SetBrightness(1, 1, 1)
				}
				
				;---------------------------------------
				;DEBUGGER LITTLE SQUARE FOR BLANK SCREEN
				;---------------------------------------
				; Gui, Show, x0 y0 w100 h100	
				
				;--------------------------------------------------------
				; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
				; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
				;--------------------------------------------------------
				IF (POWER_SCREEN_SAVE_OFF="False")
					SendMessage, 0x112, 0xF170, 2,, Program Manager
				;--------------------------------------------------------
				; CARE WITH SWITCHING OFF MONITOR SOME DUPLICATING MONITOR 
				; CAN LOSE THAT SETTING
				;--------------------------------------------------------
				
				;--------------------------------------------------------
				; how to create a black screen - Ask for Help - AutoHotkey Community
				; https://autohotkey.com/board/topic/37016-how-to-create-a-black-screen/
				; ----
				; Or you could really turn it off using:
				; SendMessage, 0x112, 0xF170, 1,, Program Manager   ; 0x112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
				;--------------------------------------------------------
				
				SetTimer,Mouse_Idle_Timer, 10 ; RAPID RESPONSE BRING OUT OF DIM
				; SET TIMER TO THE SECOND 1000MS FOR EASIER COUNTER TIME_OUT
				Mouse_Idle_Flip_Flop_Toggle := "True"
			}
		
		IF IN_DAY=TRUE
		{
			Gui, HIDE
			;--------------------------------------------------------
			; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
			; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
			;--------------------------------------------------------
			IF (POWER_SCREEN_SAVE_OFF="False")
			{
				SendMessage, 0x112, 0xF170, -1,, Program Manager
				; ---------------------------------------------------
				; POWER UP THE SCREEN AND WILL FLASH ON FOR SECOND THEN BACK TO POWERSAVE
				; SO ANSWER IS MOVE THE MOUSE TO BRING OUT OF SCREENSAVER
				; WORKER YES
				; ---------------------------------------------------
				; ACTIVATING THE SCREEN AGAIN WITH THIS COMMAND IS NOT ACTUALLY
				; REQUIRED _ IT IS SOME SYSTEM SCREEN SAVER AND MOUSE OR KEY IDLE
				; ACTIVATES IT THAT WAY
				; ---------------------------------------------------
				
				MouseGetPos, xpos, ypos
				XPOS+=1
				YPOS+=1
				MouseMove, xpos, ypos
				XPOS-=1
				YPOS-=1
				MouseMove, xpos, ypos
			}
		
			Monitor.SetBrightness(127, 127, 127)
			
			SetTimer,Mouse_Idle_Timer, 1000 
			; RAPID RESPONSE BRING OUT OF DIM
			; SET TIMER TO THE SECOND 1000MS FOR EASIER COUNTER TIME_OUT
			Mouse_Idle_Flip_Flop_Toggle := "False"
		}
	}
	O_IN_DAY_1:=%IN_DAY%

RETURN

IS_IN_DAY:

FormatTime, T,, HHmm
If ( ( T >= 0000 and T <= 0500 ) or ( T >= 2200 and T <= 2359 ) )
{
	IN_DAY=FALSE
}
ELSE
{
	IN_DAY=TRUE
}

; IN_DAY=FALSE

IF IN_DAY=TRUE 
	ALLOW_DIMMER := "False"


RETURN

; ------------------------------------------------------------------
isWindowFocusExplorer( winTitle ) {
winID := WinExist( winTitle )
If ( !winID )
	Return false
WinGetClass, this_class, ahk_id %winID%
WinActiveExplorer=false

If this_class=CabinetWClass
	WinActiveExplorer=true
If this_class=ExploreWClass
	WinActiveExplorer=true

WinGet style, MinMax, ahk_id %WinID%
;tooltip, %style%

if WinActiveExplorer=false
	Return false

;1 maximized 0 normal -1 minimized
If style=-1
{
	;SoundBeep , 2500 , 100
	Return false
}

Return true
}


; ------------------------------------------------------------------
isWindowFullScreen( winTitle ) {
; checks if the specified window is full screen

winID := WinExist( winTitle )

If ( !winID )
	Return false
	
; ONLY CHECK VALID TITLE WITH TEXT OR ELSE AFRAID CAPTURE DESKTOP 
WinGetTitle, Title_VAR, ahk_id %winID%
If ( !Title_VAR )
	Return false
	
; MSGBOX % Title_VAR
; WINDOWS XP REPORT Program Manager FOR DESKTOP
; ---------------------------------------------
IF Title_VAR=Program Manager	
	Return false

WinGet style, Style, ahk_id %WinID%
WinGetPos ,,,winW,winH, ahk_id %WinID%

; 0x800000 is WS_BORDER.
; 0x20000000 is WS_MINIMIZE.
; no border and not minimized

Return ((style & 0x20800000) 
or winH < A_ScreenHeight 
or winW < A_ScreenWidth) ? false : true

; ----
; Detect Fullscreen application? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/38882-detect-fullscreen-application/
; ----
}

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
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return


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

; ------------------------------------------------------------------
EOF:                           ; on exit
Display.OnExit()               ; release and free the library
ExitApp     
; ------------------------------------------------------------------

; ------------------------------------------------------------------
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
; ---------------------------------------------------------------------
; exit the app

; JUNK BOX FINDING
; ------------------------
; 01 OF 03
; ------------------------
; Tooltip % "You pressed "  GetKeyName(SubStr(A_ThisHotkey,2))
; ------------------------
; 02 OF 03
; ------------------------
; If (CurrentX = LastX and CurrentY = LastY) 
; {
;  Mouse_Idle += 1 ; IDLE ACTIVTY _ ACTIVE = FALSE
; }
; Else 
; {
;  Mouse_Idle = 0  ; IDLE ACTIVTY _ ACTIVE = TRUE
; }
; 
; If (Mouse_Idle > 4)
; ------------------------
; ------------------------
; 04 OF 04
; ------------------------
; global cnt := DllCall("user32.dll\GetSystemMetrics", "Int", 80)  ; get the number of display monitors on a ; desktop
; global tmp := []                                                 ; initialize an Array
;#
; loop % cnt                                                               ; loop throw all monitors
; tmp[A_Index] := Display.GetMonitorBrightness(1).CurrentBrightness        ; get the current brightness for ; all monitors
; ------------------------
; 04 OF 04 __ DETECT FULL BY SCAN ALL WINDOW IF FULLSCREEN NOT ACTIVE ONE
; ------------------------
; ----
; close all windows? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/4104-close-all-windows/
; ----
; ------------------------
; -- OF -- 
; ------------------------
; if isFullScreen = 1 ;true __ UNABLE TO USE TRUE WITH IF



; Hi, There, Room
; Worked On This One Thursday and Friday
; Brushing Up on My Technique From Newbie Entry Level at AutoHotKey
; The Other Day Had a New System, and Was Code Something Halt at Not Have Proper Registered Driver
; and Near Impossible t Check if a Driver is Registered even if Unregister Again it is Still present in ; Registry to Show
; At the End For My JoyPad Controller and Reciever 
; Rumble Pad and F710 Logitech
; I Had the Code Register the Driver Before the Next Line of Code Ready to Run _, And All Happen Seamlessly
; Now It Register the Driver Once a Day 
; Check Work_a_Round
; I Could Do to Another computer as Only Once
; But Sometimes My Reinstall is to the Same Computer Name and User
; That Was Another Subject Get to the Point
; Here is My Code
; I Found Recently Google Does Not Do a Nice Display HTML Page On Google Drive 
; But DropBox Does It Nicely
; Here The Project Code is For Dimm the Brightness of the Monitor 
; Something Better than Screen Saver Blank Switch Off
; Unfortunately
; Windows 7 Was Okay Dimmer after IDLE Time
; but Windows Above Require SECPOL.MSI Running to Set that Thing
; Which only Pro Version Pay Extra For Having to |Allow Dimmer After IDLE
; BackWards About Environmentally Friendly 
; Microsoft __ How Dimmer _ Din
; Over and Out
; ~
; Fri 15 September 2017 13:11:54----------
; ----
; Roids Polaroids Mach 2 HardCore: AutoHotKey -- 14-Brightness.ahk -- Dimmer
; http://roidsrim-minimal.blogspot.co.uk/…/AutoHotkey--14-Bri…
; ----
; ----
; Hi, Room Worked On This One Thursday and Friday... - Matthew Lancaster
; https://www.facebook.com/matthew.lancaster.4/posts/10208906605496145?comment_id=10208906608336216&comment_tracking=%7B%22tn%22%3A%22R%22%7D
; ----
; ----
; Hi, Room Worked On This One Thursday and Friday Brushing Up on My Technique Fr...
; https://plus.google.com/u/0/+RoidsRim-0/posts/LxaXsLkHD12
; ----
; 
; BLOGGER
