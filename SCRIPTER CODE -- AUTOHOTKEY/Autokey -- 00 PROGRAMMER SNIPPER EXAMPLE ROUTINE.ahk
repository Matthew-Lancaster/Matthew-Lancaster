#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------


; THE CAMERA COPY SET HERE
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
; C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 29-COPY CAMERA PHOTO IMAGES.VBS



; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path

; -------------------------------------------------------------------
; #InstallKeybdHook
; #InstallMouseHook
; -------------------------------------------------------------------


XX=0
WINWAITACTIVE ahk_class FileLocatorProMainFrame_49C8C49D-3E31-4B75-88D2-F8FC5BDCBC84
SLEEP 100
SETTIMER TIMER_MOUSE_DOWN,200
RETURN




TIMER_MOUSE_DOWN:
	XX+=1
	IF XX>1000
	{
		SETTIMER TIMER_MOUSE_DOWN,OFF
		XX=0
		SETTIMER TIMER_MOUSE_UP,200
	}
	TOOLTIP % XX ,500,100
	Click, WheelDown
	IFWINNOTACTIVE ahk_class FileLocatorProMainFrame_49C8C49D-3E31-4B75-88D2-F8FC5BDCBC84
		PAUSE
RETURN

TIMER_MOUSE_UP:
	XX+=1
	IF XX>1000
	{
		SETTIMER TIMER_MOUSE_UP,OFF
		XX=0
		SETTIMER TIMER_MOUSE_DOWN,200
	}
	TOOLTIP % XX ,500,100
	Click, WheelUp
	IFWINNOTACTIVE ahk_class FileLocatorProMainFrame_49C8C49D-3E31-4B75-88D2-F8FC5BDCBC84
		PAUSE
RETURN


; -------------------------------------------------------------------
; 0001 __ TEST ROUTINE
; -------------------------------------------------------------------
GLOBAL ArrayCount
; Each array must be initialized before use:
FN_Array_1 := []
FN_Array_2 := []
FN_Array_3 := []
DATE_MOD_Array := []

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


Loop % ArrayCount
{
	Element := FN_Array_1[A_Index]
	IF !Element
	{
		ArrayCount-=1
	}
}
MSGBOX %ArrayCount%
EXITAPP
RETURN
; -------------------------------------------------------------------





; -------------------------------------------------------------------
; 0002 __ MSGBOX
; -------------------------------------------------------------------
; 4096 BRING TO FOREGROUND
; 4100 BRING TO FOREGROUND + YES NOT
; FOREGROUND + YES NOT
; 4096 + 4
; -----------------------------------------------
; MSGBOX ,4100,,% "------ -- ------ ----- ---- ---- ---------`n`n"
; IFMSGBOX YES
; -----------------------------------------------
MSGBOX ,4096,,NOT GOODSYNC PORTABLE D-DRIVE WITH ASUS 4G AT THE MOMENT`n`nAS IMAGE AT GOOGLE PHOTO NOT IN-LINE YET
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; 0003 __  DETECTOR LAST KEYCODE AND 2ND SC CODE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #InstallKeybdHook
; #InstallMouseHook
; -------------------------------------------------------------------
; -------------------------------------------------------------------
LOOP
{
	TOOLTIP % A_PRIORKEY
}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CODE ACTIVTE RUN NOT MATTER WHERE IN SCRIPT SO REM NOT INTERFERE
; ; -------------------------------------------------------------------
; ; -------------------------------------------------------------------
; SetFormat, Integer, Hex
; Gui +ToolWindow -SysMenu +AlwaysOnTop
; Gui, Font, s14 Bold, Arial
; Gui, Add, Text, w100 h33 vSC 0x201 +Border, {SC000}
; Gui, Show,, % "// ScanCode //////////"
; LOOP
; {
; Loop 9
  ; OnMessage( 255+A_Index, "ScanCode" ) ; 0x100 to 0x108
; }

; ScanCode( wParam, lParam ) {
 ; Clipboard_TMP := "SC" SubStr((((lParam>>16) & 0xFF)+0xF000),-2) 
 ; GuiControl,, SC, %Clipboard_TMP%
; }
; RETURN
; ; -------------------------------------------------------------------




; -------------------------------------------------------------------
; SESSION 0004 __  DETECT PRESS IN TITLE BAR AREA AND INITIATE INSTRUCTION _ LIKE REMOVE APP LOAD ANOTHER
; Thu 28-May-2020 12:31:54
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SIXER PROJECT TODAY FROM THIS MORNING
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #1
; ROUTINE NAME
; SETTIMER CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__4G,6000
; FILE SCRIPTOR
; Autokey -- 58-AUTO REPEAT BROWSER FUNCTION SET.ahk
; CODE TO KNOCK OUT THE NAG SCREEN FROM FACEBOOK THAT BOLLOCKS 
; AND ROUTINE OTHER ELSEWHERE
; HARD AS CHROME DON;T DISPLAY ANY INFO ABOUT MSGBOX THING
; SO ROUTINE INSTRUCTION DELAY AND PRESS ESCAPE KEY OVER EVERYTHING IF CERTAIN PAGE 
; AND INITIATE BY FIRST SCRIPT CODE
; THEY PUT NAG SCREEN TALKER PEOPLE SPAM IT STAY FOR NEAR WHOLE DAY
; FOR DON'T LIKE SOMETHING
; ZUCK JUST GOT OUT HIS NAPPY
; Thu 28-May-2020 08:27:28
; Thu 28-May-2020 08:39:22 -- 11 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #2 
; GOODSYNC NOT REQUIRE TO RUN EVERY COMPUTER AT MOMENT ONLY THE PORTABLE VERSION D-DRIVE AND ONE COMPUTER
; SENSITIVE JOB GROUP REQUIRE LOOK AT BEFORE RUN MULTI
; Thu 28-May-2020 08:55:08
; Thu 28-May-2020 08:58:39 -- 3 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #3
; ROUTINE NAME
; RIGHT_CLICK_TO_LOAD_REAL_WINDOWS_10_CALCULATOR_NOT_WINAERO_COM_CONVERTOR_WINDOWS_7
; FILE SCRIPTOR
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; -------------------------------------------------------------------
; THIRD PROJECT THIS MORNING
; DETECT PRESS IN TITLE BAR AREA AND INITIATE INSTRUCTION _ LIKE REMOVE APP LOAD ANOTHER
; UNDERNEATH APP PUSH ONTO CONTEXT MENU UP -- NOT WANTER EXTRA HARDY
; -------------------------------------------------------------------
; ORIGINAL INTENTION CLICK TITLE BAR CLOSE APP LOAD ANOTHER
; CALC FOR WINDOWS 7 ONLY RUN BY SPECIAL APP 
; REPLACE GET WINDOWS 10 VERSION 
; ONLY BY TASK-BAR LINK 
; AND HOTKEY NOT DETECTABLE
; -------------------------------------------------------------------
; Thu 28-May-2020 09:15:12
; Thu 28-May-2020 12:10:00 -- 2 HOUR 55 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; #4 
; FILE VB CODE
; RS232_LOGGER - Microsoft Visual Basic [design]
; 4TH PROJECT TODAY OUT OF SIXER
; Sub TIMER_FRONT_DOOR_TIMER()
; WORK PROVIDE RESET METHOD EVERY 24 HOUR MIDNIGHT
; SEEM BUG SOMEWHERE -- GOOD FORM RELOAD -- ANOTHER TO DO OF THEM
; OTHER QUIT UNTIL MORNIGN CALL AGAIN
; NEW IDEA SEMI GET THERE LOADER IT ON WHEN CHROME INITIATE
; --
; 4TH PROJECT TODAY OUT OF SIXER
; DOCUMENT @ Autokey -- 00 PROGRAMMER SNIPPER EXAMPLE ROUTINE.ahk
; NEW DOCUMENT START TO PUT TOGETHER CODE SNIPPET AS WORK ALONG
; EASIER FIND AND TEST THERE
; --
; MAKE RS232 RESET MIDNIGHT
; AND NOT COMPLAIN BEEN SHUT DOWN PROPER WHEN AN EXTRA MEASURE 
; TO ENSURE SHUT DOWN PROPER OR DOOR STATE NOT TIME
; ALSO FIXX AR OTHER WHILE THERE
; A BIT BACKWARD BUTTY
; NEAR DONE WORK WRITE FOR VB2008 ALSO
; NOW TAKE COMPARE WORK GET DO
; -------------------------------------------------------------------
; Thu 28-May-2020 12:34:02
; Thu 28-May-2020 13:43:00 -- 1 HOUR 8 MINUTE
; -------------------------------------------------------------------
; #5
; -------------------------------------------------------------------
; Shell_Explorer - Microsoft Visual Basic [design] - [Form1 (Code)]
; -------------------------------------------------------------------
; 005 PROJECT TODAY
; MAKE DOWNLOAD FOLDER SHOW
; CAPITAL A LOT OF THING HARDCODER
; TIDY SORTER LINKER AND CAPITAL URL SEMI CODER
; MAKE NET 2 GO ANOTHER METHOD VIDEO
; -------------------------------------------------------------------
; Thu 28-May-2020 14:00:18
; Thu 28-May-2020 15:20:00 -- 1 HOUR 19 MINUTE
; -------------------------------------------------------------------
; --------------------------------------------------------------------
; SIXER 006 PROJECT TODAY
; --------------------------------------------------------------------
; MODIFY DATE CODER  
; MAKE DO GIVER BLUE SELECT BUTTON FOR LAST CHOOSE
; AND ALSO WRITE TO FILE
; MODIFY DATE AND ALSO RENAME A LITTLE BIT
; --------------------------------------------------------------------
; Thu 28-May-2020 15:22:11
; Thu 28-May-2020 15:48:00 -- 25 MINUTE
; --------------------------------------------------------------------
; ALL DAY CODE FROM 8 AM
; AND SHOPPING MARKET IN MORNING BEFORE
; FINE FIDDLY NUMBER TO SPEND
; 3 MONTH SINCE SHOPPING MARKET OPEN FOR DELIVERY
; CLOSE CALL BILL END OF MONTH
; STILL 2 WORK LEFT TO DO -- SEARCHER THING -- ITEMS GONE
; FALL OUT POCKET INFO
; AND LOST INFO SCREEN FROM TAB PAGE MESS AH -- HARD WORK FIND THEM
; --------------------------------------------------------------------
; PUBLISH ORGANISE TEXT INFO WORK
; --------------------------------------------------------------------
; Thu 28-May-2020 15:48:00 
; Thu 28-May-2020 16:32:00 -- 44 MINUTE
; --------------------------------------------------------------------

; HI ROOM
; GET THING FROM EBAY & ARRIVE OVER 
; SIXTIETH DAY LATE 
; BEYOND FEEDBACK RETURN
; HONG KONG
; BETTER IF NOT OVER 180 DAY OR NOT PAYPAL GUARANTEE
; IT DO COMPLAIN
; SUCCESS IF WAIT LONGER
; PREVIOUS I HAD 40 DAY LATE
; RIGHT AT BEGIN EBAY FROM HONG KONG
; & SIT IN OFFICE BUTTY JUST ARRIVE THERE MAYBE
; I WAS NEAR GO TELEPHONE EBAY 
; & ONLY 2ND CALL MAKE THIS YEAR
; & VODAFONE PUT UP PRICE SO MY DISCOUNT GONE
; THEY TALKER IS NEVER GET DROP AGAIN
; BUT TARIFF CHANGER
; & I AM WITH SUPER DISCOUNT 
; WITH TWO MOBILE
; HALF THE COST OF £30 MONTH PRESENT AR
; NOW I EVENTUALLY FORCER TO RING UP ANOTHER BLOODY AUTHORITY
; STUCK HIGH AH BILL
; DISCOUNT TAKEN OFF WITH TARIFF CHANGE 
; SOMETIME DEPEND NOT HAPPEN BEFORE GIVER
; --
; THURSDAY BEFORE 
; EXTREME WORK CODER
; 7 PROJECT UPDATE
; FEW HOUR EACH
; TIDY UP THING IMPROVE
; 1ST FACEBOOK
; GOT NEW NAG SCREEN
; BLAST AWAY
; THU 28-MAY-2020 08:27:28
; THU 28-MAY-2020 08:39:22 -- 11 MINUTE
; 2..
; GOODSYNC PORTABLE NOT REQUIRE TO RUN EVERY COMPUTER IF D-DRIVE COPY
; SENSITIVE WORK GO ON -- UNTIL DONE 
; SYNC FEW COMPUTER
; THU 28-MAY-2020 08:55:08
; THU 28-MAY-2020 08:58:39 -- 3 MINUTE
; --
; 3..
; RIGHT CLICK TO LOAD REAL WINDOWS 10 CALCULATOR NOT WINAERO_COM REPLACE WINDOWS_7
; DETECT MOUSE PRESS IN TITLE BAR AREA & INITIATE 
; LOAD WIN-10-CALC
; UNDERNEATH APP MOUSE PUSH ONTO CONTEXT MENU UP -- NOT WANTER EXTRA HARDY
; CALC FOR WINDOWS 7 ONLY RUN BY SPECIAL APP WINAERO.COM REPLACE WINDOWS 10 VER FORCE NOT UNDO
; WIN10 ONLY BY TASK-BAR LINK 
; HOTKEY TOP RIGHT CALC NOT DETECTABLE
; ALL MS-OS CALC
; THU 28-MAY-2020 09:15:12
; THU 28-MAY-2020 12:10:00 -- 2 HOUR 55 MINUTE
; --
; 4..
; VB6 CODE & 2008 LATER
; RS232_LOGGER
; TIMER_FRONT_DOOR_LOGGER
; WORK PROVIDE RESET METHOD EVERY 24 HOUR MIDNIGHT
; & ANY RESET & NOT DISRUPT IT THINK FAULT
; PROVIDE FORCE PROCESS KILL ON SPECIAL WAY
; THU 28-MAY-2020 12:34:02
; THU 28-MAY-2020 13:43:00 -- 1 HOUR 8 MINUTE
; --
; 5..
; SHELL_EXPLORER - VB6
; GOT EXPLORER SHELL MENU LIKE STARTMENU 
; & A BIT MORE HARDCORE 
; MAKE DOWNLOAD FOLDER SHOW
; CAPITAL A LOT HARDCODER
; IT GRAB URL FROM FOLDER & LINKER LIKE STARTMENU DONE AS 2ND PART
; 1ST PART LOT ITEM FILE SYSTEM -- TIDIER
; DETECT DEAD LINK
; SPECIAL TRICK LATER MAKE RUN QUICKER
; 1ST TIME ACCESS SHORTCUT INFO TAKE LONGER
; INCODE RUN IT THROUGH MANY & REST QUICKER
; REQUIRE SPEED UP -- WHAT DONE BEFORE
; THU 28-MAY-2020 14:00:18
; THU 28-MAY-2020 15:20:00 -- 1 HOUR 19 MINUTE
; --
; SIXER..
; FILE DATE CODER  VB
; MAKE DO GIVER BLUE SELECT BUTTON FOR THE LAST CHOOSE & ALSO STORE TO FILE
; MODIFY DATE CODE WITH RENAME
; THU 28-MAY-2020 15:22:11
; THU 28-MAY-2020 15:48:00 -- 25 MINUTE
; --
; DINNER
; 7..
; CLIPBOARD LOGGER 
; FEATURE SHARE CLIPPER ALL NETWORK
; DESIGN CODE BETTER INFO
; MOST THERE YET TO COMPLETE
; 12 MIDNIGHT GONE
; --
; SHOPPING MARKET HOVE FRIDAY
; BIG LOAD WALK ALL WAY & SAINSBURY WOULD
; BOOK SHOPPING FOR SUNDAY 1ST IN 3 MONTH
; LOW FLY BANK BILL 1ST
; NOT PROPER PASSCODE 
; ATM HATE TOOK MY CARD CODE ISSUE CORRECT
; ROYAL MAIL NOT DELIVER ENVELOPE
; --
; 2020-05-28 20-24-08 - SONY DSC-HX60V - DSC03026.JPG
; --
; OVER AND ON
; HA HA
; ~
; Sat 30-05-2020 10:29
; 3047 -- 5591DFB


; HI ROOM
; AND THEN AR
; AS MESSENGER POST BEFORE 
; SCREENSHOT
; IT GOT FEATURE RECENT
; WHEN SELECT FOLDER
; AND ASK 2 STEP NETWORK OPTION
; SELECT NETWORK DRIVE ALSO
; AND IT TAKE PATH REQUEST AT NETWORK LOCATION
; AND IF NETWORK DRIVE LOCATION NOT EXIST
; THERE THREE MAIN FIRST DRIVE FOR EACH NETWORK
; IF NOT EXIST IT CHECK ALL OTHER 3 NETWORK DRIVE FOR THAT COMPUTER
; IN CASE OF COORDINATION ERROR USER
; OR FORGOT WHAT DRIVE WERE
; OPERATION OF VALUE
; ----------------------------------------
; GRUB4DOS DO THAT TYPE OF THING
; CHECK PATH DRIVE UNTIL FINDER
; ----------------------------------------
; 5..
; SHELL_EXPLORER - VB6
; GOT EXPLORER SHELL MENU LIKE STARTMENU
; & A BIT MORE HARDCORE
; MAKE DOWNLOAD FOLDER SHOW
; CAPITAL A LOT HARDCODER
; IT GRAB URL FROM FOLDER & LINKER LIKE STARTMENU DONE AS 2ND PART
; 1ST PART LOT ITEM FILE SYSTEM -- TIDIER
; DETECT DEAD LINK
; SPECIAL TRICK LATER MAKE RUN QUICKER
; 1ST TIME ACCESS SHORTCUT INFO TAKE LONGER
; INCODE RUN IT THROUGH MANY & REST QUICKER
; REQUIRE SPEED UP -- WHAT DONE BEFORE
; THU 28-MAY-2020 14:00:18
; THU 28-MAY-2020 15:20:00 -- 1 HOUR 19 MINUTE
; ----------------------------------------
; HI ROOM GET THING FROM EBAY & ARRIVE OVER
; HTTP://TINYURL.COM/Y7V7VYRO
; SATURDAY, 30 MAY 2020 AT 10:39
; ----------------------------------------
; HA HA
; ----------------------------------------
; OVER AND ON AND HARDCORE
; ----------------------------------------
; CIAO
; ----------------------------------------
; ~
; Sat 30-May-2020 11:18:00
; Sat 30-May-2020 11:24:00
; Text Size  1471 -- CRC32 1B0F5590
; Text Size  1447 -- CRC32 A25E1FEC



; -------------------------------------------------------------------
; NOTHING SPECIAL -- RUN IT AND
; -------------------------------------------------------------------
; //autohotkey.com/board/topic/21105-crazy-scripting-scriptlet-to-find-scancode-of-a-key/?p=138256
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; ---- DECLARE FOR HERE ROUTINE AND SETTER GO
; -------------------------------------------------------------------
; TRIG_FIND_STATEL_WINDOWS_7=
; TRIG_FIND_STATEL_WINDOWS_7_TIMER_1=
; TRIG_FIND_STATEL_WINDOWS_7_TIMER_2=
; MOUSE_DOWN_Calculator=
; SETTIMER RIGHT_CLICK_TO_LOAD_REAL_WINDOWS_10_CALCULATOR_NOT_WINAERO_COM_CONVERTOR_WINDOWS_7,100
; RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; HAVE ROUTINE HERE -- 
; WHEN RIGHT CLICK DOWN I PRESS ON APPLICATION WITH KILL PROCESS AND UNDERNEATH DO RIGHT CONTEXT MENU UP
; -------------------------------------------------------------------
; #IfWinActive, Calculator ahk_class CalcFrame
; RButton::
; {
	; MouseGetPos, offsetx, offsety	; x x x
	; offsetx := offsetx - 0	     	; x O x  O = tip of mouse cursor
	; offsety := offsety - 0	     	;
	; IF offsetx>41
	; IF offsetx<562
	; IF offsety<38   ; --  HITT IN TITLE BAR AREA NOT ANY OTHER BUTTON THERE -- AND DEPEND SIZE BUTTON TYPE THING
	; {
	; MOUSE_DOWN_CALCULATOR=TRUE
	; }
; }
; RETURN
; #ifwinactive
; -------------------------------------------------------------------
; -------------------------------------------------------------------
RIGHT_CLICK_TO_LOAD_REAL_WINDOWS_10_CALCULATOR_NOT_WINAERO_COM_CONVERTOR_WINDOWS_7:
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	IF TRIG_FIND_STATEL_WINDOWS_7_TIMER_1<%A_NOW%
	IF TRIG_FIND_STATEL_WINDOWS_7=TRUE
	IF (stateL = "U")
	{
		TRIG_FIND_STATEL_WINDOWS_7=
		TRIG_FIND_STATEL_WINDOWS_7_TIMER_1=
		TRIG_FIND_STATEL_WINDOWS_7_TIMER_2=%A_NOW%
		TRIG_FIND_STATEL_WINDOWS_7_TIMER_2 += 2, SECONDS
		TOOLTIP
	}
	IfWinNOTActive Calculator ahk_class CalcFrame
		RETURN
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------

	MouseGetPos, offsetx, offsety	; x x x
	offsetx := offsetx - 0	     	; x O x  O = tip of mouse cursor
	offsety := offsety - 0	     	;
	GetKeyState, stateL, LButton 
	GetKeyState, stateR, RButton 

	; ---------------------------------------------------------------
	; ORIGINAL INTENTION CLICK TITLE BAR CLOSE APP LOAD ANOTHER
	; CALC FOR WINDOWS 7 ONLY RUN BY SPECIAL APP 
	; REPLACE GET WINDOWS 10 VERSION 
	; ONLY BY TASK-BAR LINK 
	; AND HOTKEY NOT DETECTABLE
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	IF offsetx>41
	IF offsetx<562
	IF offsety<38   ; --  HITT IN TITLE BAR AREA NOT ANY OTHER BUTTON THERE -- AND DEPEND SIZE BUTTON TYPE THING
	IF MOUSE_DOWN_CALCULATOR=TRUE
	{
		TOOLTIP
		MOUSE_DOWN_CALCULATOR=
		Process, CLOSE, Calc1.exe
		; Process, CLOSE, Calc1.exe
		RUN, "C:\PStart\# NOT INSTALL REQUIRED\CALC WIN 04 10\Calc Windows 10.exe"
	}
	; ---------------------------------------------------------------
	IF offsetx>41
	IF offsetx<562
	IF offsety<38   ; --  HITT IN TITLE BAR AREA NOT ANY OTHER BUTTON THERE -- AND DEPEND SIZE BUTTON TYPE THING
	IF (stateL = "D")
	IF TRIG_FIND_STATEL_WINDOWS_7_TIMER_2<%A_NOW%
	{
		CoordMode, ToolTip, Screen  ; Place ToolTips at absolute screen coordinates.
		WinGetPos TOOLTIP_X, TOOLTIP_Y,,,Calculator ahk_class CalcFrame
		TOOLTIP_X+= 8
		TOOLTIP_Y-= 70
		
		TOOLTIP % "RIGHT CLICK TO LOAD REAL WINDOWS 10 CALCULATOR`n`nNOT -- WINAERO.COM CONVERTOR WINDOWS 7",%TOOLTIP_X%,%TOOLTIP_Y%
		TRIG_FIND_STATEL_WINDOWS_7=TRUE
		TRIG_FIND_STATEL_WINDOWS_7_TIMER_1=%A_NOW%
		TRIG_FIND_STATEL_WINDOWS_7_TIMER_1 += 2, SECONDS
		; TOOLTIP % "Offset x = " . offsetx . ", y = " . offsety
	}
RETURN
; ----
; -------------------------------------------------------
; -------------------------------------------------------
; CODE NOT REQUIRE HOTKEY AND NOW USER TIMER ROUNTINE OF 
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; -------------------------------------------------------
; -------------------------------------------------------
; -------------------------------------------------------
; -------------------------------------------------------
; NOT WORK NOT KEYCODE MY COMPUTER -- CALC NOT DETECTABLE
; -------------------------------------------------------
; Launch_App2::
; RUN, "C:\PStart\# NOT INSTALL REQUIRED\CALC WIN 04 10\Calc Windows 10.exe"
; RETURN
; #ifwinactive

; NOT WORK NOT KEYCODE MY COMPUTER -- CALC NOT DETECTABLE
; -------------------------------------------------------
; sc121::
; RUN, "C:\PStart\# NOT INSTALL REQUIRED\CALC WIN 04 10\Calc Windows 10.exe"
; RETURN
; #ifwinactive
; -------------------------------------------------------
; -------------------------------------------------------
; -------------------------------------------------------
; -------------------------------------------------------
