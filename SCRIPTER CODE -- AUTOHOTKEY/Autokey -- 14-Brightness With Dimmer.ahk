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

; -------------------------------------------------------------------
; SESSION PLENTY
; -------------------------------------------------------------------
; Sun 26-Jan-2020 01:40:28
; Sun 26-Jan-2020 02:50:14 -- 1 HOUR 10 MINUTE
; -------------------------------------------------------------------
;
; -------------------------------------------------------------------
; HAD TO REDO WORK OTHER DAY -- TIMNER AND HOUR STAY AWAKE 
; DURING NIGHT TIME ADDITION SCREEN SAVER AFTER 10 PM O'CLOCK
;
; AND THEN MERGE COMPARE
; SOME DID AR SYNC LOST IT
; WORK WAS MOVE ADD HOUR ON FILE TO DATA FOLDER RATHER THAN HERE
; NEXT TO SCRIPTOR FOLDER
; -------------------------------------------------------------------


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

;		IF (A_ComputerName="8-MSI-GP62M-7RD")
;			DIMMER_ONLY_NOT_BLANK := "True"
;       CEHCKER FIND LINE ABOVE IS REQUIRE
; Sun 26-Jan-2020 02:49:24

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


; ---------------------------------------------------------------
; REQUIRE GIVE BEFORE INCLUDE
; ---------------------------------------------------------------
; SET THE FLAG TO BEGIN TIMER
; ---------------------------------------------------------------
ADD_MINUTE_BEFORE_SCREEN_SAVER=
; CARRY ON THE TIMER SET
ADD_MINUTE_SCREEN_SAVER=
; ---------------------------------------------------------------

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 04 of 04_SETTIMER.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER\Class_Monitor_Master\SRC\Class_Monitor.ahk     ; include the class here


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
; SESSION 001
; Fri 15 September 2017 01:09:04---- ABOVE WORK BEGINNER
; ------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 002
; Wed 21 March     2018 02:10:21---- Work to Detect Explorer Was Active ; As it Remains Present When Closed and Minimal Sometimes Error There 
; had to be Resolved _ Explorer and Other Program make Quicker Timeout ; Brightness Change _ Few Hour-ing
; ------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 003
; WORK EXTRA INCLUDE DAY TIME MODE BRIGHTNESS UP
; -------------------------------------------------------------------
; FROM TIME __ Wed 20-Jun-2018 15:20:52
; TO   TIME __ Wed 20-Jun-2018 15:54:00
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 004
; WORK _ INCLUDE A BLACK SCREEN ALSO WITH DIM _ SLEEP BETTER AT NIGHT IN ROOM
; -------------------------------------------------------------------
; FROM TIME __ Tue 03-Jul-2018 22:01:03
; TO   TIME __ Wed 04-Jul-2018 01:24:25 4 HOUR 24 MIN FOR A SMALL 
;                                       PIECE OF CODE BOO 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 005
; WORK _ ADD DISPLAY OFF WORKING POWER ENERGY SAVER COMBINED WITH BLANK SCREEN
; -------------------------------------------------------------------
; FROM TIME __ Wed 04-Jul-2018 07:46:36 __ SESSION 01 OF 02
; TO   TIME __ Wed 04-Jul-2018 08:30:44

; FROM TIME __ Wed 04-Jul-2018 17:42:15 __ SESSION 02 OF 02
; TO   TIME __ Wed 04-Jul-2018 18:19:19
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 00*
; MORE WORK _ DEBUG NOT WORK CORRECTLY TIMER AND LATCH 
; -------------------------------------------------------------------
; FROM TIME __ Wed 04-Jul-2018 22:35:01 _ DEBUG VARIABLE SYNTAX
; TO   TIME __ Wed 04-Jul-2018 23:22:00 _ 50 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 007
; MORE WORK _ THE LAST OF THE LATCHING BUGGS SORTED
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 10:06:11 _ DEBUG VARIABLE SYNTAX
; TO   TIME __ Thu 05-Jul-2018 10:24:00 _ 18 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 008
; MORE WORK _ MY 7-ASUS HAS A DUPLICATE MONITOR SYSTEM AND MONITOR 
; POWER LEAVES DUPLICATE MODE SO ONLY BLANK SCREEN USER
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 22:22:00
; TO   TIME __ Thu 05-Jul-2018 22:25:00 _ 3 MINUTE 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 009
; MORE WORK _ FINIAL LOOK AT CODE PROBLEM MONITOR WASN'T COMING OUT OF 
; STANDBY AT THE PRESET TIME IN MORNING _ ANSWER WAS TO MOVE A THE MOUSE
; A LITTLE BIT-PIXEL AND BACK AGAIN REMARK COMMENTS ARE NEAR CODE
; -------------------------------------------------------------------
; FROM TIME __ Thu 05-Jul-2018 22:20:55
; TO   TIME __ Thu 05-Jul-2018 12:02:00 _ 1 HOUR 42 MINUTE
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 010
; -------------------------------------------------------------------
; ADD THE RS232 PIR CONTROL 
; -------------------------------------------------------------------
; TOOK BIT OF STRESS INTRO AND THEN EASIER THAN THOUGHT 
; LAST TOOK BIT OF CLEAN TIDIER UP SORT A FEW BUG SEEM TO BE ON
; JITTER BUG
; TWO SESSION EARLY HOUR FROM 1 AM TO 5 AM AND THEN MORNING AFTER WAKE AGAIN
; EARLIER TO POST DELIVERY
; PIR FILENAME AND (COMPUTERNAME) FILE PROVIDE FROM RS232
; -------------------------------------------------------------------
; THE PIR CONTROL SIGNAL NOT IN HERE CODE -- OVER THERE
; EXAMPLE FILENAME
; \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer #NFS.txt
; D:\VB6\VB-NT\00_Best_VB_01\RS232 LOGGER PIR\RS232 LOGGER PIR.exe
; WHEN MY MSI COMPUTER DOWNSTAIR_S IS DONE HERE
; Mon 20-Jan-2020 18:04:40
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; FROM TIME __ Sat 13-Apr-2019 01:16:44
; TO   TIME __ Sat 13-Apr-2019 05:14:58 _ 5 HOUR
; FROM TIME __ Sat 13-Apr-2019 09:08:08
; TO   TIME __ Sat 13-Apr-2019 11:33:55 _ 2 HOUR 15 MINUTE
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 011
; -------------------------------------------------------------------
; NEW CODER AS HERE
; THIS CODE IS IN REPEAT BLOCK USE 4 TIME
; -------------------------------------------------------------------
; WANT REAL MOUSE TIME __ TimeIdle
; ONLY WANT PSYCHICAL KEYBOARD TIME __ TimeIdle
; [ Tuesday 09:51:00 Am_14 May 2019 ]
; -------------------------------------------------------------------
; THAT IS COMPLETE 
; THE SITUATION IS 
; KEYBOARD TIMEIDLE IS ONLY WANT PSYCHICAL AS THERE ARE USE SOME 
; F5 REFRESH PAGE ON THE INTERNET AND AS THIS IS A GLORY HOLE SCREEN SAVER
; WE DON'T WANT THEM SIMULATED KEY PRESS AND THE KEYBOARD HOOK IS INSTALLED FOR THAT
; BUT WITH MOUSE WE WANT OPPOSITE THAT SIMULATED MOUSE MOUSE IS ABLE BRING 
; OUT OF SCREEN SAVER
; GET MY DRIFTER
; HAD TO GO TO BED FALLER ALSLEEP ON THIS ONE AND WAKE TO THE GO
; [ Tuesday 10:12:30 Am_14 May 2019 ]
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Tue 14-May-2019 02:20:42 -- NIGHT BEFORE STARTER
; TO   TIME __ Tue 14-May-2019 02:20:42
; FROM TIME __ Tue 14-May-2019 09:18:45 -- MORNING
; TO   TIME __ Tue 14-May-2019 10:15:54
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 012
; -------------------------------------------------------------------
; BANGED OUT THE CODE 
; TO
; I GOT RS232 PIR FOR SCREEN SAVER
; AND I GOT SOME CODE WHICH HAS ARTIFICIAL KEY-PRESS TO KEEP THING RUNNER
; AS YOU LEARN SCREEN SAVER ALWAYS COMES OUT OF IT BY A KEY PRESS EVEN 
; IF ARTIFICIAL 
; THE BEST WAY AROUND THAT _ TOOK HARD WORK
; IS WHEN THE SCREEN SAVER COME ON QUICKLY PUT BACK
; I REQUIRE _ PRIOR-KEY SPECIAL VARIABLE
; I REQUIRE _ HOTKEY ON THE F5 KEY TO DETERMINE WITH KEY-STATE IS ARTIFICIAL OR NOT
; I REQUIRE _ SETLEVEL FROM ON SEPARATE CODE TO MAKE PRIOR-KEY WORK IN ANOTHER SCRIPT
; I REQUIRE _ USE OF THE HOTKEY TO RECORD THE TIME OF KEY-PRESS
; I REQUIRE _ MAYBE REQUIRE MORE OTHER HOTKEY AS USE TAB FOR ANOTHER THING BUT LATER TO THEM
; I REQUIRE _ ALL THE PROPER DELAY EXTENDED OF MOUSE MOVE AND F5 REFRESH HOTKEY HAS EXPIRED TIMER
; I REQUIRE _ AND USE A DELAY FOR PIR IN CODE DIFFERENT LONGER FOR MAIN WORK COMPUTER ANY COMPUTER NAME SET TO WHAT WANT AND SHORTER IF IN PIR OFF TO ACTIVATE BLANK SCREEN MODE
; I REQUIRE _ AFTER THE SCREEN SAVER HAPPEN AND THE DELAY SINCE HOTKEY F5 ALSO REQUIRE THE VARIABLE IS RESET AFTER BEEN USE TO PUT INTO SCREEN SAVER BLANK SCREEN
; THAT WRAPPED IT UP
; AND IT DETECT AND ARRAY OF WINDOW TITLE THAT COMMONLY USER AS REFRESH TYPE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Sat 08-Jun-2019 18:40:43 _ CLIPBOARD COUNT _ 384
; TO   TIME __ Sat 08-Jun-2019 23:09:10 _ CLIPBOARD COUNT _ 557
; TO   TIME __ Sat 08-Jun-2019 23:09:10 _ Result: 2 hours, 28 minutes and 27 seconds
; -------------------------------------------------------------------
; Sat 08-Jun-2019 23:09:10
; Sat 08-Jun-2019 18:40:43

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

 


#InstallKeybdHook
;#InstallMouseHook

; IF YOU GOT ANYTHING KEYBOARD HOOKING SENDING A KEY OR MOUSE
; AND USE HERE A_TimeIdlePhysical __ INSTEAD OF A_TimeIdle
; ONLY Physical INPUT NOT SIMULATED SENDER

; -------------------------------------------------------------------

; TRICK PART OF CODE IS MANY DON'T TELL HAVE TO FIND IDLE AS WELL ALSO THE ACTIVE PART
; ALL VERY WELL SWITCH THE SCREEN SAVER ON BUT____ AND THEN OFF SCREEN SAVER


OnExit, EOF                         ; set onexit to the label "EOF"
Display := New Monitor()            ; initialize / start the class

; SCRIPT ============================================================

DetectHiddenWindows, on

; PAUSE



CoordMode, Mouse, Screen
Mouse_Idle = 0 
Mouse_Idle_Flip_Flop_Toggle := "True"
SCREEN_SAVER_NOT_ACTIVE := "TRUE"
OLD_SCREEN_SAVER_NOT_ACTIVE=
SCREEN_SAVER_NOT_ACTIVE_TIME_DELAY=
LastX = 0
LastY = 0
VAR_A__TimeIdle_1_OF_4:=0
VAR_A__TimeIdle_2_OF_4:=0
VAR_A__TimeIdle_3_OF_4:=0
VAR_A__TimeIdle_4_OF_4:=0
VAR_Z__TimeIdle_1 = 4000
VAR_Z__TimeIdle_4_DEFAULT = 80000 ; 4 MINUTE
VAR_Z__TimeIdle_3_FORCE = 2000 
VAR_Z__TimeIdle_2 = %VAR_Z__TimeIdle_2_DEFAULT%
VAR_Z__TimeIdle = %VAR_Z__TimeIdle_1%
OLDWinActive = 0
WinActive_2 = 0

USE_A_TimeIdlePhysical=FALSE
IF (A_ComputerName="7-ASUS-GL522VW")
	USE_A_TimeIdlePhysical=TRUE
IF (A_ComputerName="8-MSI-GP62M-7RD")
	USE_A_TimeIdlePhysical=TRUE
IF (A_ComputerName="5-ASUS-P2520LA") 
	USE_A_TimeIdlePhysical=TRUE

; -------------------------------------------------------------------
; WANT REAL MOUSE TIME __ TimeIdle
; ONLY WANT PSYCHICAL KEYBOARD TIME __ TimeIdle
; [ Tuesday 09:51:00 Am_14 May 2019 ]
; -------------------------------------------------------------------
; THAT IS COMPLETE 
; THE SITUATION IS 
; KEYBOARD TIMEIDLE IS ONLY WANT PSYCHICAL AS THERE ARE USE SOME 
; F5 REFRESH PAGE ON THE INTERNET AND AS THIS IS A GLORY HOLE SCREEN SAVER
; WE DON'T WANT THEM SIMULATED KEY PRESS AND THE KEYBOARD HOOK IS INSTALLED FOR THAT
; BUT WITH MOUSE WE WANT OPPOSITE THAT SIMULATED MOUSE MOUSE IS ABLE BRING 
; OUT OF SCREEN SAVER
; GET MY DRIFTER
; HAD TO GO TO BED FALLER ALSLEEP ON THIS ONE AND WAKE TO THE GO
; [ Tuesday 10:12:30 Am_14 May 2019 ]
; -------------------------------------------------------------------

IF USE_A_TimeIdlePhysical=FALSE
	VAR_A__TimeIdle_1_OF_4=%A_TimeIdle%
ELSE
{
	VAR_A__TimeIdle_1_OF_4=%A_TimeIdlePhysical%
	IF A_TimeIdleMouse<%VAR_A__TimeIdle_1_OF_4% 
	VAR_A__TimeIdle_1_OF_4=%A_TimeIdleMouse%
}
	

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

BLANK_DIMMER_TIME=60*2

BLANK_DIMMER_TIME=40
; IF (A_ComputerName="4-ASUS-GL522VW")    ; --- TRY HERE SCREEN GO TO QUICKER BLANK DARK POWER GONE -- 
	; BLANK_DIMMER_TIME=60*5
IF (A_ComputerName="8-MSI-GP62M-7RD")
	BLANK_DIMMER_TIME=60*1

	
BLANK_DIMMER=%A_Now%
BLANK_DIMMER+= %BLANK_DIMMER_TIME%, Seconds

SoundBeep , 1000 , 100
SoundBeep , 3000 , 100

; IF (A_ComputerName="1-ASUS-X5DIJ")
	; PAUSE
; IF (A_ComputerName="2-ASUS-EEE")
	; PAUSE
; IF (A_ComputerName="4-ASUS-GL522VW")
	; PAUSE
	

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

SoundBeep , 1000 , 100
SoundBeep , 3000 , 100

IF OSVER_N_VAR=5 ; XP
	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	
	
	
Gui, Color, black
Gui +AlwaysOnTop
Gui -Caption
; Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%	
Gui, hide
; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
SendMessage, 0x112, 0xF170, 0,, Program Manager

GOSUB, MONITOR_BRIGHTNESS_UP

SETTIMER,MOUSE_IDLE_TIMER, 1000     ; CHECK EVERY SECOND
SETTIMER TIMER_PREVIOUS_INSTANCE,1
SETTIMER,CHECK_FILENAME_1_HOUR,1000

; -------------------------------------------------------------------
FN_ARRAY_AUTO_KEY := SET_ARRAY_AUTO_KEY()
ARTIFICIAL_F5=-2
ARTIFICIAL_F5_A_Now:=A_Now
ARTIFICIAL_F5_A_Now-=40

RS232_LOGGER_PIR_VAR=0
OLD_RS232_LOGGER_PIR_VAR=-1
IF (A_ComputerName="4-ASUS-GL522VW")
	SetTimer,RS232_LOGGER_TIMER_RUN_EXE, 4000

IF (A_ComputerName="1-ASUS-X5DIJ")
{
	RS232_IDLE_SET_DELAY_1=20000
	RS232_IDLE_SET_DELAY_2=1
}
IF (A_ComputerName="2-ASUS-EEE")
{
	RS232_IDLE_SET_DELAY_1=20000
	RS232_IDLE_SET_DELAY_2=1
}
IF (A_ComputerName="4-ASUS-GL522VW")
{
	RS232_IDLE_SET_DELAY_1=40000
	RS232_IDLE_SET_DELAY_2=1
}

IF (A_ComputerName="1-ASUS-X5DIJ")
	SetTimer,RS232_LOGGER_TIMER_CHANGE, 1000
IF (A_ComputerName="2-ASUS-EEE")
	SetTimer,RS232_LOGGER_TIMER_CHANGE, 1000
IF (A_ComputerName="4-ASUS-GL522VW")
	SetTimer,RS232_LOGGER_TIMER_CHANGE, 1000
; -------------------------------------------------------------------
; -------------------------------------------------------------------
RETURN


; -------------------------------------------------------------------
; INITIALIZE
; -------------------------------------------------------------------
; 1ST SUBROUTINE / FUNCTION
; -------------------------------------------------------------------

; Loop % FN_Array_1.MaxIndex()
; {
	; Element := FN_Array_1[A_Index]
	; ; MSGBOX %Element%
	; IfWinExist, %Element%
		; XR_2=1
		; XR_4=%Element%
; }
; FN_ARRAY_AUTO_KEY := SET_ARRAY_AUTO_KEY()
SET_ARRAY_AUTO_KEY() {
	SET_ARRAY_AUTO_KEY := []
	GLOBAL ArrayCount
	ArrayCount := 0
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Your notifications - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook | Error - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Privacy error - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook - Mozilla Firefox"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Facebook - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Deborah Hall -"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Dibs Dabs -"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="She - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Follow the Sun - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Hallelujah - YouTube - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Rain Alarm - Google Chrome"
	ArrayCount += 1
	SET_ARRAY_AUTO_KEY[ArrayCount]:="Rain Alarm - Mozilla Firefox"
RETURN SET_ARRAY_AUTO_KEY
}



RS232_LOGGER_TIMER_RUN_EXE:
	
	FN_VAR_RS232_FILENAME:="D:\VB6\VB-NT\00_Best_VB_01\RS232 LOGGER PIR\RS232 LOGGER PIR_APP_RUNNER_TICKER_#NFS_EX.TXT"
	IFEXIST, %FN_VAR_RS232_FILENAME%
	{
		FileGetTime, A_Script_MODIIFED_DATE , %FN_VAR_RS232_FILENAME% 
		; DIFFICULT IF FILENAME NOT GIVEN CORRECT %VAR% OR 
		; WITHOUT QUOTE -- RETURN CURRENT TIME
		; -----------------------------------------------------------
		A_Script_MODIIFED_DATE+= 15, Seconds
		IF A_Script_MODIIFED_DATE>%A_NOW%
			RETURN
	}
	
	; ---------------------------------------------------------------
	; AS HIDDEN WINDOW HARD TO FIND -- NOW MAKE CHAIN DOG TRIGGER
	; ---------------------------------------------------------------
	; NOT WITH GOOD FIND REPEAT LOAD THIS CODE BEEN GO LONG TIME 
	; UNTIL FAULT FINDER NOW
	; Sun 09-Feb-2020 20:05:33
	; Sun 09-Feb-2020 20:51:04 -- 46 MINUTE
	; ---------------------------------------------------------------
	FN_VAR_RS232_EXE:="D:\VB6\VB-NT\00_Best_VB_01\RS232 LOGGER PIR\RS232 LOGGER PIR.exe"
	IfWinNotActive RS232_LOGGER - Microsoft Visual Basic [ ahk_class wndclass_desked_gsk
	IFWINNOTEXIST RS232_LOGGER ahk_class ThunderFormDC
	IFWINNOTEXIST RS232_LOGGER ahk_exe RS232 LOGGER PIR.exe
	IFEXIST, %FN_VAR_RS232_EXE%
	{
		Run, %FN_VAR_RS232_EXE%,,HIDE
	}
	; IFNOTEXIST, %FN_VAR_RS232_EXE%
	; {
		; MSGBOX NOT EXIST`n%FN_VAR_RS232_EXE%
	; }
RETURN 

RS232_SUB: ; ----
RS232_LOGGER_TIMER_CHANGE: ; ----

	IF ADD_MINUTE_SCREEN_SAVER
	IF ADD_MINUTE_SCREEN_SAVER>%A_NOW%
		RETURN

	; HERE COME FROM
	; "D:\VB6\VB-NT\00_Best_VB_01\RS232 LOGGER PIR\RS232 LOGGER.vbp"
	; ---------------------------------------------------------------
	FN_VAR_TXT=C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER #NFS__%A_ComputerName%.txt

	; \\2-ASUS-EEE\2_ASUS_EEE_01_c_drive\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER #NFS__2-ASUS-EEE.TXT

	IFNOTEXIST, %FN_VAR_TXT%
	{
		RS232_LOGGER_PIR_VAR=0
	}
	ELSE
	{
		; WANT ON -------------------------------------------------------
		RS232_LOGGER_PIR_VAR=1
	}
	

	; TOOLTIP % RS232_LOGGER_PIR_VAR
	; TOOLTIP %A_TimeIdle% " -- " %RS232_IDLE_SET_DELAY_1%
	
	IF OLD_RS232_LOGGER_PIR_VAR=%RS232_LOGGER_PIR_VAR%
	IF RS232_LOGGER_PIR_VAR=1
		RETURN

	OLD_RS232_LOGGER_PIR_VAR=%RS232_LOGGER_PIR_VAR%
	
	; WANT ON -------------------------------------------------------
	IF RS232_LOGGER_PIR_VAR=1
	{
		MouseMove, 10, 10, , R
		MouseMove, -10, -10, , R
		SoundBeep , 2500 , 100
		Monitor.SetBrightness(127, 127, 127)
	}

	
	; -------------------------------------------------------------------
	; WANT REAL MOUSE TIME __ TimeIdle
	; ONLY WANT PSYCHICAL KEYBOARD TIME __ TimeIdle
	; [ Tuesday 09:51:00 Am_14 May 2019 ]
	; -------------------------------------------------------------------
	; THAT IS COMPLETE 
	; THE SITUATION IS 
	; KEYBOARD TIMEIDLE IS ONLY WANT PSYCHICAL AS THERE ARE USE SOME 
	; F5 REFRESH PAGE ON THE INTERNET AND AS THIS IS A GLORY HOLE SCREEN SAVER
	; WE DON'T WANT THEM SIMULATED KEY PRESS AND THE KEYBOARD HOOK IS INSTALLED FOR THAT
	; BUT WITH MOUSE WE WANT OPPOSITE THAT SIMULATED MOUSE MOUSE IS ABLE BRING 
	; OUT OF SCREEN SAVER
	; GET MY DRIFTER
	; HAD TO GO TO BED FALLER ASLEEP ON THIS ONE AND WAKE TO THE GO
	; [ Tuesday 10:12:30 Am_14 May 2019 ]
	; -------------------------------------------------------------------

	IF USE_A_TimeIdlePhysical=FALSE
		VAR_A__TimeIdle_2_OF_4=%A_TimeIdle%
	ELSE
	{
		VAR_A__TimeIdle_2_OF_4=%A_TimeIdlePhysical%
		IF A_TimeIdleMouse<%VAR_A__TimeIdle_2_OF_4% 
			VAR_A__TimeIdle_2_OF_4=%A_TimeIdleMouse%
	}
	
	QUICKER_OFF=FALSE
	
	; ---------------------------------------------------------------
	; A_PriorKey WILL NOT WORK FOR WHAT WANT HERE
	; UNLESS OTHER CODE Autokey -- 58-Auto Repeat Browser Function Set.ahk
	; HAS HERE
	; 
	; SendLevel 1
	; ---------------------------------------------------------------
	
	ARTIFICIAL_IDLE=%A_Now%
	ARTIFICIAL_IDLE-=%ARTIFICIAL_F5_A_Now%,s
	
	; TOOLTIP % ARTIFICIAL_IDLE
	; IF ARTIFICIAL_IDLE<8  ; ---- SECOND THAT OUGHT TO DO HER

	IF A_PriorKey<>F5
		ARTIFICIAL_F5=1
	
	IF A_PriorKey=F5
	IF ARTIFICIAL_F5=0    ; ---- 0 = ARTIFICIAL ---- 1 = PSYCHICAL
	Loop % FN_ARRAY_AUTO_KEY.MaxIndex()
	{
		Element := FN_ARRAY_AUTO_KEY[A_Index]
		IfWinActive, %Element%
			QUICKER_OFF=TRUE
	}

	; ---------------------------------------------------------------
	; TOOLTIP % ARTIFICIAL_IDLE " -- " ARTIFICIAL_F5 " -- " QUICKER_OFF
	; DONE END HERE THIS LINE
	; [ Saturday 22:59:20 Pm_08 June 2019 ]
	; ---------------------------------------------------------------
	
	IF QUICKER_OFF=FALSE
	IF VAR_A__TimeIdle_2_OF_4 < %RS232_IDLE_SET_DELAY_1%
		RETURN
	IF QUICKER_OFF=TRUE
	IF VAR_A__TimeIdle_2_OF_4 < %RS232_IDLE_SET_DELAY_2%
		RETURN
	

	; WANT OFF ------------------------------------------------------
	IF RS232_LOGGER_PIR_VAR=0
	{
		
		; TOOLTIP %A_NOW% 

		; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
		; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
		SendMessage, 0x112, 0xF170, 2,, Program Manager
		SoundBeep , 2500 , 100
		ARTIFICIAL_F5=1
	}

RETURN


~$F5:: ; ---- 
{
ARTIFICIAL_F5:=GetKeyState("F5","P")
ARTIFICIAL_F5_A_Now:=A_Now
}
RETURN



; -------------------------------------------------------------------
; WORK CODER TO COMPLETE HERE
; FINALLY AT 
; Fri 03-Jan-2020 03:19:44
; FROM 
; Fri 03-Jan-2020 02:37:19
; -------------------------------------------------------------------
; FEW THING FINDER
; THE MENU CODE SET FILE 02 OF 03
; HOLD THE VARIABLE FOR MENU SELECT
; AND UNABLE HAVE CODE IN THERE AS REACH END NOT A RETURN
; SO CODE FOR SUBROUTINE GO IN MENU FILE 03 OF 03
; IF CODE ARE LOCATE ON 02 OF 03
; THEY RUN TWICE
; -------------------------------------------------------------------
; AND CODER ADD
; THAT WHEN THE TIMER VARIABLE DO 
; IT ALSO TRIGGER THAT SCREEN SAVER SWITCH IS SHOW VISIBLE
; -------------------------------------------------------------------

SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND:
	EXIST_VAR_SCREEN_SAVER=
	IF ADD_MINUTE_BEFORE_SCREEN_SAVER
		EXIST_VAR_SCREEN_SAVER=TRUE

	FILENAME_VB=C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer ADD_SOME_#NFS_EX_%A_ComputerName%.TXT
	IF FILEEXIST(FILENAME_VB)
		EXIST_VAR_SCREEN_SAVER=TRUE
	
	IF EXIST_VAR_SCREEN_SAVER
	{
		ADD_MINUTE_BEFORE_SCREEN_SAVER=
		ADD_MINUTE_SCREEN_SAVER=%A_Now%
		ADD_MINUTE_SCREEN_SAVER+=1 , Hours
		GOSUB SCREEN_SAVER_TO_SHOW_SCREEN
		if FileExist(FILENAME_VB)
			FileDelete, % FILENAME_VB
		SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		Gui, -caption +toolwindow +AlwaysOnTop +Disabled -SysMenu +Owner  ; +Owner avoids a taskbar button.
		Gui Color, White
		Gui font, s30 bold, Arial
		Gui, Add, Text,, Screen Saver Set On
		; Gui, Show, , Title of Window 
		Gui, Show, NoActivate, Title of Window  ; NoActivate avoids deactivating the currently active window.
		SLEEP 4000
		Gui, Hide
	}
RETURN


SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND_1X_2E_COMPUTERNAME:

	FILENAME_VB=C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER ADD_ACTIVE_1X_2E_#NFS_EX_%A_ComputerName%.TXT
	IF FILEEXIST(FILENAME_VB)
		EXIST_VAR_SCREEN_SAVER=TRUE
	
	IF EXIST_VAR_SCREEN_SAVER
	{
		ADD_MINUTE_BEFORE_SCREEN_SAVER=
		ADD_MINUTE_SCREEN_SAVER=%A_Now%
		ADD_MINUTE_SCREEN_SAVER+=10 , Minutes
		GOSUB SCREEN_SAVER_TO_SHOW_SCREEN
		if FileExist(FILENAME_VB)
			FileDelete, % FILENAME_VB
	}
RETURN

SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND_3L_5P_8M_COMPUTERNAME:

	FILENAME_VB=C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER ADD_ACTIVE_3L_5P_8M_#NFS_EX_%A_ComputerName%.TXT
	IF FILEEXIST(FILENAME_VB)
		EXIST_VAR_SCREEN_SAVER=TRUE
	
	IF EXIST_VAR_SCREEN_SAVER
	{
		ADD_MINUTE_BEFORE_SCREEN_SAVER=
		ADD_MINUTE_SCREEN_SAVER=%A_Now%
		ADD_MINUTE_SCREEN_SAVER+=10 , Minutes
		GOSUB SCREEN_SAVER_TO_SHOW_SCREEN
		if FileExist(FILENAME_VB)
			FileDelete, % FILENAME_VB
	}
RETURN

CHECK_FILENAME_1_HOUR: ; ----
	; FROM MENU OPTION ----------------------------------------------
	GOSUB SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND
	GOSUB SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND_1X_2E_COMPUTERNAME
	GOSUB SET_LOADER_FILE_FOR_SCREEN_SAVER_EXTEND_3L_5P_8M_COMPUTERNAME

	; TOOLTIP % SCREEN_SAVER_NOT_ACTIVE
	
	IF SCREEN_SAVER_NOT_ACTIVE_TIME_DELAY<%A_NOW%
		OLD_SCREEN_SAVER_NOT_ACTIVE=0
	
	IF SCREEN_SAVER_NOT_ACTIVE<>%OLD_SCREEN_SAVER_NOT_ACTIVE%
	{
		OLD_SCREEN_SAVER_NOT_ACTIVE=%SCREEN_SAVER_NOT_ACTIVE%
		IF SCREEN_SAVER_NOT_ACTIVE=FALSE
		{
			SCREEN_SAVER_NOT_ACTIVE_TIME_DELAY=%A_NOW%
			SCREEN_SAVER_NOT_ACTIVE_TIME_DELAY+= 4, Minutes
			
			IF (A_ComputerName="4-ASUS-GL522VW")
				GOSUB WRITE_FILE_SCREEN_BRIGHT_ACTIVE_COMPUTER_NAME_DOWN
			IF (A_ComputerName="8-MSI-GP62M-7RD")
				GOSUB WRITE_FILE_SCREEN_BRIGHT_ACTIVE_COMPUTER_NAME_DOWN
			IF (A_ComputerName="7-ASUS-GL522VW")
				GOSUB WRITE_FILE_SCREEN_BRIGHT_ACTIVE_COMPUTER_NAME_UP
		}
	}
	
	
RETURN


WRITE_FILE_SCREEN_BRIGHT_ACTIVE_COMPUTER_NAME_DOWN:
		
		FILENAME_2__=_01_c_drive\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER ADD_ACTIVE_1X_2E_#NFS_EX_
		
		ArrayCount = 0
		Loop, Read, C:\NETWORK_COMPUTER_NAME.txt 
		{
			NET_PATH:=A_LoopReadLine
			
			SET_GO=FALSE
			IF INSTR(NET_PATH,"1-ASUS-X5DIJ")
				SET_GO=TRUE
			IF INSTR(NET_PATH,"2-ASUS-EEE")
				SET_GO=TRUE
			IF INSTR(NET_PATH,"4-ASUS-GL522VW")
				SET_GO=TRUE
			IF INSTR(NET_PATH,"8-MSI-GP62M-7RD")
				SET_GO=TRUE
				
			IF SET_GO=TRUE
			{
				; TOOLTIP % NET_PATH
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
				TestString := "This is a test string.`r`n"  
				file.Write(TestString)
				file.Close()
			}
		}
RETURN

WRITE_FILE_SCREEN_BRIGHT_ACTIVE_COMPUTER_NAME_UP:
		
		FILENAME_2__=_01_c_drive\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-BRIGHTNESS WITH DIMMER ADD_ACTIVE_3L_5P_8M_#NFS_EX_
		
		ArrayCount = 0
		Loop, Read, C:\NETWORK_COMPUTER_NAME.txt 
		{
			NET_PATH:=A_LoopReadLine
			
			SET_GO=FALSE
			IF INSTR(NET_PATH,"3-LINDA-PC")
				SET_GO=TRUE
			IF INSTR(NET_PATH,"5-ASUS-P2520LA")
				SET_GO=TRUE
			IF INSTR(NET_PATH,"7-ASUS-GL522VW")
				SET_GO=TRUE

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
				TestString := "This is a test string.`r`n"  
				file.Write(TestString)
				file.Close()
			}
		}

RETURN






; ------------------------------------
Mouse_Idle_Timer: ; ----

COUNT_TICK_TIME=% 1000*60*24
; TOOLTIP %A_TICKCOUNT% __ %COUNT_TICK_TIME%

IF A_TICKCOUNT< %COUNT_TICK_TIME%
	RETURN

GOSUB, MONITOR_BRIGHTNESS_DIMMER_PER_DAY

DetectHiddenWindows, on
ALLOW_DIMMER := "True"

WinGetCLASS, CLASS, A
WinGetTITLE, TITLE_VAR_2, A

IF INSTR(CLASS,"MediaPlayerClassicW")
	ALLOW_DIMMER := "False"

; IF INSTR(CLASS,"Notepad++")
; ALLOW_DIMMER := "False"

IF INSTR(TITLE_VAR_2,"tube - Google Chrome")
	ALLOW_DIMMER := "False"
	
IF INSTR(TITLE_VAR_2,"YouTube - Google Chrome")
	ALLOW_DIMMER := "False"

IF INSTR(TITLE_VAR_2,"bunker.com")
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
;IF (A_ComputerName="8-MSI-GP62M-7RD")
;	ALLOW_DIMMER := "False"
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
SET_GO=TRUE
IF (ALLOW_DIMMER = "False")
	SET_GO=FALSE
	

; -------------------------------------------------------------------
; WANT REAL MOUSE TIME __ TimeIdle
; ONLY WANT PSYCHICAL KEYBOARD TIME __ TimeIdle
; [ Tuesday 09:51:00 Am_14 May 2019 ]
; -------------------------------------------------------------------
; THAT IS COMPLETE 
; THE SITUATION IS 
; KEYBOARD TIMEIDLE IS ONLY WANT PSYCHICAL AS THERE ARE USE SOME 
; F5 REFRESH PAGE ON THE INTERNET AND AS THIS IS A GLORY HOLE SCREEN SAVER
; WE DON'T WANT THEM SIMULATED KEY PRESS AND THE KEYBOARD HOOK IS INSTALLED FOR THAT
; BUT WITH MOUSE WE WANT OPPOSITE THAT SIMULATED MOUSE MOUSE IS ABLE BRING 
; OUT OF SCREEN SAVER
; GET MY DRIFTER
; HAD TO GO TO BED FALLER ALSLEEP ON THIS ONE AND WAKE TO THE GO
; [ Tuesday 10:12:30 Am_14 May 2019 ]
; -------------------------------------------------------------------

; TOOLTIP %A_TimeIdle% " -- " %A_TimeIdleMouse% " -- " %A_TimeIdlePhysical%

IF USE_A_TimeIdlePhysical=FALSE
	VAR_A__TimeIdle_3_OF_4=%A_TimeIdle%
ELSE
{
	VAR_A__TimeIdle_3_OF_4=%A_TimeIdlePhysical%
	IF A_TimeIdleMouse<%VAR_A__TimeIdle_3_OF_4% 
		VAR_A__TimeIdle_3_OF_4=%A_TimeIdleMouse%
}	

	
If VAR_A__TimeIdle_3_OF_4 < %VAR_Z__TimeIdle%
	SET_GO=FALSE
	
IF SET_GO=TRUE
{
	GOSUB, MONITOR_BRIGHTNESS_DIM
}
VAR_Z__TimeIdle=%VAR_Z__TimeIdle_2%
GOSUB, Keyboard_Idle_Timer
;#-------------------------------
Return ; End of Mouse_Idle_Timer

; ------------------------------------------------------------------
Keyboard_Idle_Timer: ; ----
; Tooltip % + A_TimeIdle " -- " VAR_A__TimeIdle 
;TEST DEBUG ___________



; TOOLTIP %A_TimeIdle% " -- " %VAR_A__TimeIdle_1_OF_4%

; -------------------------------------------------------------------
; WANT REAL MOUSE TIME __ TimeIdle
; ONLY WANT PSYCHICAL KEYBOARD TIME __ TimeIdle
; [ Tuesday 09:51:00 Am_14 May 2019 ]
; -------------------------------------------------------------------
; THAT IS COMPLETE 
; THE SITUATION IS 
; KEYBOARD TIMEIDLE IS ONLY WANT PSYCHICAL AS THERE ARE USE SOME 
; F5 REFRESH PAGE ON THE INTERNET AND AS THIS IS A GLORY HOLE SCREEN SAVER
; WE DON'T WANT THEM SIMULATED KEY PRESS AND THE KEYBOARD HOOK IS INSTALLED FOR THAT
; BUT WITH MOUSE WE WANT OPPOSITE THAT SIMULATED MOUSE MOUSE IS ABLE BRING 
; OUT OF SCREEN SAVER
; GET MY DRIFTER
; HAD TO GO TO BED FALLER ALSLEEP ON THIS ONE AND WAKE TO THE GO
; [ Tuesday 10:12:30 Am_14 May 2019 ]
; -------------------------------------------------------------------

IF USE_A_TimeIdlePhysical=FALSE
	VAR_A__TimeIdle_4_OF_4=%A_TimeIdle%
ELSE
{
	VAR_A__TimeIdle_4_OF_4=%A_TimeIdlePhysical%
	IF A_TimeIdleMouse<%VAR_A__TimeIdle_4_OF_4% 
		VAR_A__TimeIdle_4_OF_4=%A_TimeIdleMouse%
}
	
IF VAR_A__TimeIdle_4_OF_4 < %VAR_A__TimeIdle_1_OF_4%
{
	;SoundBeep , 2500 , 100
	;TEST DEBUG ___________
	GOSUB, MONITOR_BRIGHTNESS_UP

	BLANK_DIMMER=%A_Now%
	BLANK_DIMMER+= %BLANK_DIMMER_TIME%, Seconds

}
VAR_A__TimeIdle_1_OF_4=%VAR_A__TimeIdle_4_OF_4%

; A_TimeIdle - SHOW TIME SINCE LAST KEYBOARD OR MOUSE IN MILLISECOND
; THE DETECT IS IF LOWER THAN
; ---------------------------

RETURN



; ------------------------------------------------------------------
MONITOR_BRIGHTNESS_DIM: ; ----



IF ADD_MINUTE_SCREEN_SAVER
IF ADD_MINUTE_SCREEN_SAVER>%A_NOW%
	RETURN

IF (ALLOW_DIMMER = "False")
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
	SCREEN_SAVER_NOT_ACTIVE := "TRUE"

}
RETURN

; ------------------------------------------------------------------
MONITOR_BRIGHTNESS_UP: ; ----
; FIX BIG DIM SUPPOSED TO COME OUT OF DIM BUT WAS HAPPEN EVERY OTHER ODD ALTERNATE
; DON'T REQUIRE THE IF AROUND HERE
; [ Tuesday 02:20:40 Am_14 May 2019 ]
; -------------------------------------------------------------------
; NEXT BUG RS232 PIR SCREEN SAVER IS NOT WORK DURING DAY
; [ Tuesday 02:20:40 Am_14 May 2019 ]
; -------------------------------------------------------------------
; If (Mouse_Idle_Flip_Flop_Toggle = "True")
; {
	;SoundBeep , 2500 , 100
	; BRIGHTNESS UP 50% IS 100% __ Full Bright
	Monitor.SetBrightness(127, 127, 127)
	SetTimer,Mouse_Idle_Timer, 1000
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
	SendMessage, 0x112, 0xF170, 0,, Program Manager
	; SendMessage, 0x112, 0xF170, -1,, Program Manager

	Gui, HIDE
	Mouse_Idle_Flip_Flop_Toggle := "False"
	SCREEN_SAVER_NOT_ACTIVE := "FALSE"
; }
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
	

	; IF ADD_MINUTE_SCREEN_SAVER
	; IF ADD_MINUTE_SCREEN_SAVER<%A_NOW%
	IF (O_IN_DAY_1<>%IN_DAY% or SET_GO="TRUE")
	{
	
		POWER_SCREEN_SAVE_OFF := "False"
		; IF (A_ComputerName="7-ASUS-GL522VW")
			; POWER_SCREEN_SAVE_OFF := "True"
			
		; TEMP REM OUT -- SEE
		; IF (A_ComputerName="8-MSI-GP62M-7RD")
			; POWER_SCREEN_SAVE_OFF := "True"

		; POWER_SCREEN_SAVE_OFF := "False"
		
		DIMMER_ONLY_NOT_BLANK := "False"
		IF (A_ComputerName="8-MSI-GP62M-7RD")
			DIMMER_ONLY_NOT_BLANK := "True"
	
		IF ADD_MINUTE_SCREEN_SAVER
		IF ADD_MINUTE_SCREEN_SAVER<%A_NOW%
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
				SCREEN_SAVER_NOT_ACTIVE := "TRUE"
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
			SCREEN_SAVER_NOT_ACTIVE := "FALSE"
		}
	}
	O_IN_DAY_1:=%IN_DAY%

RETURN


SCREEN_SAVER_TO_SHOW_SCREEN: ; ----

	Gui, HIDE
	;--------------------------------------------------------
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER,  2 = Monitor Off
	; 0x112 = WM_SYSCOMMAND, 0xF170 = SC_MONITORPOWER, -1 = Monitor Power
	;--------------------------------------------------------
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

	Monitor.SetBrightness(127, 127, 127)
	
	SetTimer,Mouse_Idle_Timer, 1000 
	; ---------------------------------------------------------------
	; RAPID RESPONSE BRING OUT OF DIM
	; SET TIMER TO THE SECOND 1000MS FOR EASIER COUNTER TIME_OUT
	; ---------------------------------------------------------------
	Mouse_Idle_Flip_Flop_Toggle := "False"
	SCREEN_SAVER_NOT_ACTIVE := "FALSE"
RETURN


IS_IN_DAY: ; ----

FormatTime, T,, HHmm
If ( ( T >= 0000 and T <= 0500 ) or ( T >= 2200 and T <= 2359 ) )
{
	IN_DAY=FALSE
}
ELSE
{
	IN_DAY=TRUE
}

; -------------------------------------------------------------------
; TEST DEBUG PUT LIGHT OUT
; -------------------------------------------------------------------
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





#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk


TIMER_PREVIOUS_INSTANCE: ; ----
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

; ------------------------------------------------------------------
EOF: ; ----                    ; on exit
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
