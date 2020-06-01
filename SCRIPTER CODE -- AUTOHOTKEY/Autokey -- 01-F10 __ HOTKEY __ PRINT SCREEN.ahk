 ;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
;# __ 
;# __ Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 001
; ---------------------------------------------------------------
; CREATION 
; ---------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 002
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
; #Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 003
; -------------------------------------------------------------------
; GOSUB CHECK_ESC_KEY
; -------------------------------------------------------------------
; ONLY USER UP NOT DOWN
; UP ON OWN NOT WORK -- REQUIRE A HOTKEY DOWN STATE ALSO
; BEEN THROUGH BEFORE -- *~ESC UP::
; RESULT AT LAST USE GetKeyState, state, ESC  -- UP THING A ME JIG
; NOT DOUBEL HITT
; HARD THINK ANOTHER WAY TO DO THAT THING
; SURE REPEATER KEY NOT WORK IF LOOK LAST KEY
; -------------------------------------------------------------------
; A TREAT CODER TO DO AFTER EVERY THING ELSE TODAY WAS ADD ROUTINE TO 
; ADD AN OPTION TO TOGGLE THE "SEARCH RESULTS WINDOW NOTEPAD-PLUS-PLUS
; TOOK A SEARCH TO FINDER -- CREDIT LINK BELOW
; -------------------------------------------------------------------
; Tue 17-Sep-2019 15:25:01
; Tue 17-Sep-2019 18:05:00 -- 2 HOUR
; -------------------------------------------------------------------
; -------------------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
; -------------------------------------------------------------------
;WANT COPY TEXT AND REPEAT PASTE IT DOWN A LINE HOME DOWN PASTE PUT REMARK IN
; -------------------------------------------------------------------
;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk 
; HTTP://TINYURL.COM/R9OEH4F
;# ------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 004
; -------------------------------------------------------------------
; SUB -- KILL_ALL_MPC_EXE_NAME:
; WORKER AND ROUTINE THAT CALL BOTH THEM
; ESC KEY AND MOUSE LEFT TO CORNER COUPLE TIME TO TRIG
; NOW DO KILL ALL MPC PROCESS
; -------------------------------------------------------------------
; 
; Thu 30-Jan-2020 23:03:29
; Thu 30-Jan-2020 23:38:00 -- 35 MINUTE
; -------------------------------------------------------------------



; -------------------------------------------------------------------
#SingleInstance force
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off
DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path

Send {shift up}





VAR_REPEAT_F5_TOOGLE=
SETTIMER REPEAT_F5_BASHING,20000

SETTIMER TIMER_PREVIOUS_INSTANCE,1

GLOBAL VAR_COUNTER

GLOBAL FILE_SCRIPT_COUNT
GLOBAL FILE_SCRIPT
GLOBAL XPOS
GLOBAL YPOS

; GLOBAL O_ID
O_ID=0

GLOBAL OLD_id
GLOBAL OLD_Title_VAR
GLOBAL OLD_STATE_CAP
GLOBAL OutputVar_4

GLOBAL O_OutputVar_store
O_OutputVar_store=

GLOBAL PART_RENAME_VAR

SETTIMER TIMER_ENTER,OFF

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; REQUIRE HOOK METHOD WHEN USER ELSEWHERE
; -------------------------------------------------------------------
; Thu 28-May-2020 10:35:00
; -------------------------------------------------------------------
; #InstallKeybdHook
; #InstallMouseHook
; -------------------------------------------------------------------
; TWO ROUTINE TO CHECK HOTKEY
; -------------------------------------------------------------------
IF TRUE=FALSE
	{
	LOOP
	{
		TOOLTIP % A_PRIORKEY
	}
	RETURN
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------
IF TRUE=FALSE
	{
	SetFormat, Integer, Hex
	Gui +ToolWindow -SysMenu +AlwaysOnTop
	Gui, Font, s14 Bold, Arial
	Gui, Add, Text, w100 h33 vSC 0x201 +Border, {SC000}
	Gui, Show,, % "// ScanCode //////////"
	LOOP
	{
	Loop 9
	  OnMessage( 255+A_Index, "ScanCode" ) ; 0x100 to 0x108
	}

	ScanCode( wParam, lParam ) {
	 Clipboard_TMP := "SC" SubStr((((lParam>>16) & 0xFF)+0xF000),-2) 
	 GuiControl,, SC, %Clipboard_TMP%
	}
	RETURN
}
; -------------------------------------------------------------------



; ----
; Crazy Scripting : Scriptlet to find Scancode of a Key - Scripts and Functions - AutoHotkey Community 
; https://autohotkey.com/board/topic/21105-crazy-scripting-scriptlet-to-find-scancode-of-a-key/
; ----

; HERE THE FUNCTION ROUTINE FOR GOODSYNC
; --------------------------------------
GOSUB F5_ROUTINE
; --------------------------------------

GLOBAL STATE_XYPOS_COUNTER
GLOBAL OLD_STATE_XYPOS_COUNTER
STATE_XYPOS_COUNTER=0
OLD_STATE_XYPOS_COUNTER=0
OLD_HWND_4=0
SETTIMER TIMER_TOP_LEFT_MOUSE_CLOSE_MPC,10

GLOBAL STATE_XYPOS_COUNTER
GLOBAL OLD_STATE_XYPOS_COUNTER
STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH=0
OLD_STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH=0
OLD_HWND_5=0
SETTIMER TIMER_TOP_LEFT_MOUSE_REFRESH_BROWSER,10

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

MESSENGER_KEY=

; -------------------------------------------------------------------
O_ID=0
OLD_id=0
OLD_Title_VAR=0
OLD_STATE_CAP=0
OutputVar_4=0
; -------------------------------------------------------------------
SETTIMER WINDOW_CHECK_IF_WANT_PUT_CAPS_LOCK_OFF_OR_ON,100

; -------------------------------------------------------------------
WSCRIPT_FOCUS_SET_FLAG_01=
WSCRIPT_FOCUS_SET_FLAG_02=
; -------------------------------------------------------------------
SETTIMER TIMER_WSCRIPT_FOCUS_LEFT_KILL,1000

ESCAPE_KEY_COUNT=0


; -------------------------------------------------------------------
; -------------------------------------------------------------------
; THE CODE START TO ENTER ROUTINE FUNCTION THAT USE RETURN ON FROM 
; HERE
; NOT FOR SURE TRUE THAT INITIAL CODE END AS 
; FUNCTION SET WITH RETURN AFTER 
; THE MAIN RETURN COULD COME HERE
; MAYBE A TIDIER -- CHECK OUT LATER
; Tue 17-Sep-2019 18:15:00
; -------------------------------------------------------------------
; -------------------------------------------------------------------

#IfWinActive, Microsoft Visual Basic [design]
F5::SEND ^{f5}
RETURN
#ifwinactive

; -------------------------------------------------------------------
; Difference between IfWinActive and #IfWinActive - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/78534-difference-between-ifwinactive-and-ifwinactive/
; ----
; -------------------------------------------------------------------
; WHAT HAPPEN HERE
; -------------------------------------------------------------------
; ANYCODE WANT HOTKEY AND THE LINE IS AFTER ONE ABOVE
; WONT WORK
; UNLESS BETWEEN THERE IS HERE LINE #ifwinactive
; WHEN THE LINE HERE IT NOT STOP THE HOTKEY BEFORE
; AND THEN POSITION YOUR HOTKEY IN FRONT OF ROUTINE WANT USER
; -------------------------------------------------------------------
; Tue 17-Sep-2019 16:02:42
; -------------------------------------------------------------------
; IGNORE ABOVE THIS LINE IS NOT LOCATABLE ANYWHERE IN CODE 
; MUST BE RUN OVER LIKE DECLARE
; -------------------------------------------------------------------
; Wed 18-Sep-2019 23:01:31
; -------------------------------------------------------------------
	
*~ESC::
	GOSUB CHECK_ESC_KEY
RETURN	
#ifwinactive




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



; ----
; How do I make a mute button - Ask for Help - AutoHotkey Community 
; https://autohotkey.com/board/topic/85600-how-do-i-make-a-mute-button/
; ----

#IFWINNOTACTIVE ahk_class wndclass_desked_gsk
$F1::Send {VOLUME_MUTE}	
RETURN
#ifwinactive

; TIMER_GET_HWND:
	; WinGet, OLD_HWND, ID, ahk_class Winamp Gen

	
#IfWinActive ahk_class Chrome_WidgetWin_1
~LButton::
IfWinNotExist ClipBoard Logger
	RUN D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe
RETURN
#ifwinactive
	
; -------------------------------------------------------------------
; ANYCODE WANT HOTKEY AND THE LINE IS AFTER ONE ABOVE
; WONT WORK
; UNLESS BETWEEN THERE IS HERE LINE #ifwinactive
; -------------------------------------------------------------------


	
	
; HOT STRING -- EASY ENOUGH -- BUT WHEN TYPE HIMA MUST FOLLOW BUT SPACE OR RETURN 
; must type an ending character after typing btw, such as Space, ., or Enter).

; MESSENGER_KEY=Hi Marianne and Eddie 
; THIS ONE ENDED -- Mon 26-Aug-2019 11:01:30

; MESSENGER_KEY=HI MARIANNE;  & EDDIE


:*:hima::
MESSENGER_KEY=HI MARIANNE
GOSUB STRING_INVERT_MESSENGER
SENDINPUT %MESSENGER_KEY%
RETURN





; :*:22::
; SENDINPUT 2008{TAB}05{TAB}01{TAB}10  ; -- ABBY
; SENDINPUT 1982{TAB}02{TAB}01{TAB}10  ; -- SEKA
; SENDINPUT 1995{TAB}04{TAB}01{TAB}10    ; -- URAN
RETURN


; ADDITIONAL NOTE ABOUT BUG
; THIS CODE GOOD FOR THE N ; ' T -- WOUDLN'T BUT BUG 
; THE METHOD HAS TYPE WHILE YOU TYPE SO NEXT THING
; IT NOT SMOOTH 
; THERE IS A CODE I SAW THAT PREVENT TYPE HAPPEN WHILE 
; THE SENDINPUT FOR HOTSTRING HAPPEN
; AND I WILL GET ONTO THAT SHORTER
; [ Thursday 04:38:00 Pm_13 June 2019 ]

; -------------------------------------------------------------------
; Setting a Semicolon as a Hotkey? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/3423-setting-a-semicolon-as-a-hotkey/
; ----
; SEMICOLON
; You can also escape it with an accent:
; -------------------------------------------------------------------
; I KEEP SPELLER __ DON;T AND WOULDN;T __ RATHER THAN
;                __ DON'T AND WOULDN'T __
; -------------------------------------------------------------------
; LAZY KEYBOARD
; -------------------------------------------------------------------
; [ Tuesday 07:21:30 Am_11 June 2019 ]
; [ Tuesday 07:38:00 Am_11 June 2019 ]
; [ Tuesday 08:17:00 Am_11 June 2019 ]
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; THIS METHOD DO THE CASE INSENSITIVE PROPER
; ADD THEN ASTERISK AND NONE REQUIRE SPACE BEFORE OR AFTER 
; ASTERISK STILL REQUIRE SPACE BEFORE
; TRY ?
; USE BOTH LIKE HERE
; -------------------------------------------------------------------
:*?:n`t::SENDINPUT n't
RETURN



; -------------------------------------------------------------------
; MPC HAVE HOTKEY T TO SEND RECYCLE BIN
; -------------------------------------------------------------------
; THE CODE FOR THAT IS SOMEWHERE 
; ANOTHER SCRIPT PERHAPS
; HAVE TO PRESS THE T IN THE PLAYLIST
; -------------------------------------------------------------------
; Mon 30-Mar-2020 09:55:21
; -------------------------------------------------------------------




; -------------------------------------------------------------------
; -------------------------------------------------------------------
; NOT TOTAL OF TRUE THAT INITIAL CODE END AS 
; FUNCTION SET WITH RETURN BEFORE
; MAYBE A TIDIER
; -------------------------------------------------------------------
; -------------------------------------------------------------------




; -------------------------------------------------------------------
; WHEN REQUIRE A PROGRAM LEFT IN DEBUGGAR IDE MODE
; HERE IS WAY TO BLOCK CTRLBREAK
; THE $ AND LATER * MEAN DO ALLOW PASS THROUGH
; Sun 26-Jan-2020 04:11:00
; -------------------------------------------------------------------
#IFWINNOTACTIVE ahk_class wndclass_desked_gsk
$*CtrlBreak::RETURN
#ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; $*^C::RETURN
; #ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; $*+F5::RETURN
; #ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; CtrlBreak::RETURN
; #ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; Ctrl::RETURN
; #ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; Ctrl::MSGBOX "HH"
; #ifwinactive

; #IFWINNOTACTIVE ahk_class wndclass_desked_gsk
; ^C::RETURN
; #ifwinactive
	
; CtrlBreak::
; IFWINNOTACTIVE ahk_class wndclass_desked_gsk
    ; return  ; i.e. do nothing, which causes Control-P to do nothing in Notepad.
; Send CtrlBreak
; return


RETURN



; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CODE START -- ROUTINE SET
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; WORK TIME 
; -------------------------------------------------------------------
; Tue 14-Jan-2020 08:32:29 -- FIRST FIND IDEA
; Tue 14-Jan-2020 09:13:43 -- 45 MINUTE

; Tue 14-Jan-2020 13:17:56 -- PUBLISH ON-LINE
; Tue 14-Jan-2020 13:23:59 -- 10 MINUTE +- BIT

; LAST SESSION
; Tue 14-Jan-2020 14:34:19 -- 55 MINUTE -- TOTAL 3 SESSION 2 HOUR
; Tue 14-Jan-2020 15:30:00 -- STUCK ON A LITTLE BIT 
;                          -- AND FIRST INSTANCE 
;                          -- IMPROVE WANT NOT INCLUDE SIG LINE
;                          -- DON'T FORGET THE OBVIOUS 
;                          -- SIG LINE REQUIRE REVERSE SEARCH WAY
; -------------------------------------------------------------------


#IfWinActive .PS1 - Notepad++ ahk_class Notepad++
{
F5:: 
	SENDINPUT ^S
	SOUNDBEEP 1000,100
	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	RETURN
}
#ifwinactive





; -------------------------------------------------------------------
; SELECT 1000 PHOTO WITH HOTKEY H
; -------------------------------------------------------------------
#IfWinActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
*~H::
{
	; ControlGet, hWnd, hWnd,, Intermediate D3D Window1, A
	; MSGBOX %  hWnd
	; IF !hWnd
	; {
		TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT=0
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01,350
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,100
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,OFF

	; }
}
RETURN
#ifwinactive
; -------------------------------------------------------------------
TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01:

	IF TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT>10010
	{
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01,OFF
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,OFF
		RETURN
	}
	IF GetKeyState("LButton","P")
	{
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01,OFF
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,OFF
		RETURN
	}
	SetTitleMatchMode 2  ; PARTIAL PATH
	IfWinNOTActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
	{
		WinActivate Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
		RETURN
	}
	SENDINPUT {LEFT}
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,ON
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01,OFF
RETURN
; -------------------------------------------------------------------
TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02:
	SetTitleMatchMode 2  ; PARTIAL PATH
	IfWinNOTActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
	{
		WinActivate Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
		RETURN
	}
	SENDINPUT X
	Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
	TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT+=1
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_02,OFF
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_01,ON
RETURN


; -------------------------------------------------------------------
; -------------------------------------------------------------------
; PAGE DOWN REPEATER PHOTO WITH HOTKEY CONTROL H SHIFT H
; -------------------------------------------------------------------
#IfWinActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
*~L::    ; L FOR LATITUDE
{
	TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT=0
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01,800
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,200
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
}
RETURN
#ifwinactive
; -------------------------------------------------------------------
TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01:

	SENDINPUT {PgDn}
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01,OFF
	SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,ON
RETURN

TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02:

	IF GetKeyState("LButton","P")
	{
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		RETURN
	}
	SetTitleMatchMode 2  ; PARTIAL PATH
	IfWinNOTActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
	{
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		RETURN
	}

	IF TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT>2000
	{
		SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		RETURN
	}
	TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT+=1
	SEND ^{END}
RETURN



; -------------------------------------------------------------------
; -------------------------------------------------------------------
; PAGE DOWN REPEATER PHOTO WITH HOTKEY CONTROL H SHIFT H
; -------------------------------------------------------------------
; #IfWinActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
; *~+H::
	; TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT=0
	; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01,100
	; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,1
	; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
; RETURN
; #ifwinactive
; ; -------------------------------------------------------------------
; TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01:

	; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_01,OFF
	; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,ON
	; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
; RETURN

; TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02:

	; IF GetKeyState("LButton","P")
	; {
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		; RETURN
	; }
	; SetTitleMatchMode 2  ; PARTIAL PATH
	; If WinNOTActive Google Photos - Google Chrome ahk_class Chrome_WidgetWin_1
	; {
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		; RETURN
	; }

	; IF TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_COUNT>200
	; {
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		; SETTIMER TIMER_SELECT_20_PHOTO_WITH_HOTKEY_H_ENDER_02,OFF
		; RETURN
	; }

	; SENDINPUT {WheelDown}
; RETURN


Return


; #IfWinActive ahk_class Chrome_WidgetWin_1
; ^F5:: ; CTRL+F5 
	
	; IF !VAR_REPEAT_F5_TOOGLE
	; {
		; VAR_REPEAT_F5_TOOGLE=1
		; SETTIMER REPEAT_F5_BASHING,20000
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		; RETURN
	; }
	
	; IF VAR_REPEAT_F5_TOOGLE
	; {
		; VAR_REPEAT_F5_TOOGLE=
		; SETTIMER REPEAT_F5_BASHING,OFF
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
	; }
; Return
; #ifwinactive

REPEAT_F5_BASHING:
	IF A_TimeIdleMouse>20000
	IfWinEXIST Secure online banking home - Google Chrome ahk_class Chrome_WidgetWin_1
	{
		WinActivate
		SLEEP 500
		; IfWinActive Secure online banking home - Google Chrome ahk_class Chrome_WidgetWin_1
		SENDINPUT {F5}
		; Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	}
RETURN


; UCASE UPPER

; -------------------------------------------------------------------
; ---- THE CASE CHANGER BLOCK CODE ----------------------------------
; -------------------------------------------------------------------
#IfWinNOTActive ahk_class Notepad++
+^k:: ; SHIFT+CTRL+K ---- Converts Text To Capitalized
	VAR_INDEX=1
	GOSUB HOT_KEY_CONVERT_TExT 
RETURN
#ifwinactive
; -------------------------------------------------------------------
;#IfWinNOTActive ahk_class HwndWrapper[FreemakeVC.exe;;361378b1-207c-4a9b-9511-a53a32655ad9]
^l:: ; CTRL+L ---- Converts Text To Lower
	VAR_INDEX=2
	GOSUB HOT_KEY_CONVERT_TEXT 
RETURN
; #ifwinactive
; -------------------------------------------------------------------
#IfWinActive ahk_class Chrome_WidgetWin_1
+u:: ; SHIFT+U ---- CONVERTS TEXT TO UPPER  -- ; SHIFT NOT CONTROL AS LATER 
                                               ; HAS OPEN SOURCE OF PAGE 
											   ; IN A NEW TAB OF BROWSER
	VAR_INDEX=3
	GOSUB HOT_KEY_CONVERT_TEXT
RETURN
#ifwinactive
; -------------------------------------------------------------------
#IfWinActive ahk_class Chrome_WidgetWin_1
^u:: ; SHIFT+U ---- CONVERTS TEXT TO UPPER  -- ; SHIFT NOT CONTROL AS LATER 
                                               ; HAS OPEN SOURCE OF PAGE 
											   ; IN A NEW TAB OF BROWSER
	VAR_INDEX=3
	GOSUB HOT_KEY_CONVERT_TEXT
RETURN
#ifwinactive
; -------------------------------------------------------------------
#IfWinNOTActive ahk_class Chrome_WidgetWin_1
^u:: ; CTRL+U ---- CONVERTS TEXT TO UPPER  --- ; NOW CONTROL AND SHIFT
                                               ; CONTROL 
											   ; IF NOT BROWSER PAGE CHROME
											   ; USER SHIFT AND OR CONTROL
	VAR_INDEX=3
	GOSUB HOT_KEY_CONVERT_TEXT
RETURN
#ifwinactive
; -------------------------------------------------------------------
#IfWinNOTActive ahk_class Chrome_WidgetWin_1
+u:: ; CTRL+U ---- CONVERTS TEXT TO UPPER  --- ; NOW CONTROL AND SHIFT
                                               ; CONTROL 
											   ; IF NOT BROWSER PAGE CHROME
											   ; USER SHIFT AND OR CONTROL
	VAR_INDEX=3
	GOSUB HOT_KEY_CONVERT_TEXT
RETURN
#ifwinactive

+^D:: ; SHIFT+CTRL+D ---- CONVERTS TEXT TO DASH
	VAR_INDEX=4
	GOSUB HOT_KEY_CONVERT_TEXT
RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------




; -------------------------------------------------------------------
HOT_KEY_CONVERT_TExT:
	AutoTrim, Off
	Clipper_GET:=Clipboard
	IF !Clipper_GET
	{
		AutoTrim, On ; ---- DEFAULT
		RETURN
	}
	Clipper_1_GET=%Clipper_GET%
	Clipper_2_GET=
	StringGetPos, StrGetPos_Clipper, Clipper_GET, ~, R 
	; ---------------------------------------------------------------
	; TALK -1 IF NONE ---- StrGetPos_Clipper>0
	; Wed 22-Jan-2020 01:48:54
	; ---------------------------------------------------------------
	IF StrGetPos_Clipper>0
	{
		Clipper_1_GET:=substr(Clipper_GET, 1, StrGetPos_Clipper-1)
		Clipper_2_GET:=substr(Clipper_GET, StrGetPos_Clipper)
	}
	IF VAR_INDEX=1
		StringUpper Clipper_1_GET, Clipper_1_GET, T ; Title mode conversion
	IF VAR_INDEX=2
		StringLower Clipper_1_GET, Clipper_1_GET
	IF VAR_INDEX=3
		StringUpper Clipper_1_GET, Clipper_1_GET
	IF VAR_INDEX=4
	{
		; -----------------------------------------------------------
		; HOW CENSOR TEXT
		; CONVERT STRING REPLACE WITH DASH -- VALUE LENGTH COMPARE DUPLICATE
		; Sun 10-May-2020 11:15:46
		; CONTROL SHIFT D
		; ALSO SPACE NOT CONVERT DASH WHILE STRING ARE
		; Mon 18-May-2020 17:22:00
		; -----------------------------------------------------------
		StringLen, Length, Clipper_1_GET
		Clipper_2_GET=
		DASH_STRING:="-"
		SPACE_STRING:=" "
		LOOP, %LENGTH%
		{
		IF SUBSTR(Clipper_1_GET,A_INDEX,1)=" "
			Clipper_2_GET=%Clipper_2_GET%%SPACE_STRING%
		ELSE
			Clipper_2_GET=%Clipper_2_GET%%DASH_STRING%
		}
		Clipper_1_GET=%Clipper_2_GET%
		Clipper_2_GET=
	}
	
	Clipper_GET=%Clipper_1_GET%%Clipper_2_GET%
	; MSGBOX %Clipper_GET%
	Clipboard=%Clipper_GET%
	ClipWait
RETURN
; -------------------------------------------------------------------
; NOTE INFO 
; ----
; AHK GET CLIPBOARD STRIP LEAD SPACE - Google Search 
; https://www.google.com/search?q=AHK+GET+CLIPBOARD+STRIP+LEAD+SPACE
; ----
; AHK is Removing Leading Spaces before sending text to Clipboard - AutoHotkey Community 
; https://www.autohotkey.com/boards/viewtopic.php?t=37625
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; AutoHotkey Tip of the Week: Instant Upper Case, Lower Case, and Initial Cap Text—September 2, 2019 | Jack's AutoHotkey Blog 
; https://jacksautohotkeyblog.wordpress.com/2019/09/02/autohotkey-tip-of-the-week-instant-upper-case-lower-case-and-initial-cap-text-september-2-2019/
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Location On-Line
;--------------------------------------------------------------------
; Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk 
; HTTP://TINYURL.COM/R9OEH4F
; -------------------------------------------------------------------
; -------------------------------------------------------------------






WINDOW_CHECK_IF_WANT_PUT_CAPS_LOCK_OFF_OR_ON:

	id := WinExist("A")
	WinGetTitle, Title_VAR, ahk_id %id%
	; WinGetCLASS, CLASS_VAR, ahk_id %id%
	
	SET_GO_CAP_PUTTER=FALSE
	IF OLD_id<>%id% 
		SET_GO_CAP_PUTTER=TRUE
	IF Title_VAR<>%OLD_Title_VAR%
		SET_GO_CAP_PUTTER=TRUE
	
	IF SET_GO_CAP_PUTTER=TRUE
	{
		SetTitleMatchMode 3  ; Specify Full path
		IfWinActive mysms - Google Chrome ahk_class Chrome_WidgetWin_1
		{
			SetCapsLockState ,ON
		}
		IfWinActive ahk_class Notepad++
		{
			SetCapsLockState ,ON
		}
		SetTitleMatchMode 2  ; PARTIAL PATH
		IfWinActive eBay
		IfWinActive  - Google Chrome
		{
			SetCapsLockState ,ON
		}
		IfWinActive Your notifications
		IfWinActive  - Google Chrome
		{
			SetCapsLockState ,OFF
		}
		IfWinActive Matthew Lancaster
		IfWinActive  - Google Chrome
		{
			SetCapsLockState ,ON
		}
		IfWinActive Facebook
		IfWinActive  - Google Chrome
		{
			SetCapsLockState ,ON
		}
		IfWinActive New Tab - Google Chrome
		IfWinActive  - Google Chrome
		{
			SetCapsLockState ,ON
		}
		IfWinActive Clip Type Form
		{
			SetCapsLockState ,ON
		}
		
		IfWinActive ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009} ; GOODSYNC2GO
		{
			SetCapsLockState ,ON
		}
		IfWinActive ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A} ; GOODSYNC
		{
			SetCapsLockState ,ON
		}
		
		
		IfWinActive ahk_class CabinetWClass ; GOODSYNC
		{
			SetCapsLockState ,ON
		}
		
				
		STATE_CAP := GetKeyState("CapsLock", "T") ; True if CapsLock is ON, false otherwise.
		IF OLD_STATE_CAP<>%STATE_CAP%
		{
		IF STATE_CAP=1
			SOUNDBEEP 4000,200
		IF STATE_CAP=0
			SOUNDBEEP 1000,200
		}
		
		OLD_STATE_CAP=%STATE_CAP%
		
	}
	OLD_id=%id%
	OLD_Title_VAR=%Title_VAR%

RETURN





; -------------------------------------------------------------------
; THIS METHOD USE THE KEYBOARD CONTROL KEY NUMBER
; UNABLE TO MAKE WORKER
; PROBABLY AS CONTROL KEY CODE ARE DIFFERENT EVERY KEYBOARD MODEL
; -------------------------------------------------------------------
; ::N&SC027::
; MESSENGER_KEY=n'
; GOSUB STRING_INVERT_MESSENGER
; SENDINPUT %MESSENGER_KEY%
; RETURN


; -------------------------------------------------------------------
; HOTKEY ABLE TO DO ONE LINER BUT HOTSTRING NOT
; NOT WORKER
; -------------------------------------------------------------------
; ::nt::SENDINPUT n't  
; -------------------------------------------------------------------


; THIS METHOD NOT THE CASE PROPER
; -------------------------------------------------------------------
; ::n`;t::n't
; ::+n`;+t::N'T

; THIS METHOD DO THE CASE PROPER
; TRY THIS METHOD DONE MY HEAD IN
; -------------------------------------------------------------------
; ::n`;t::
; IF GetKeyState("Capslock", "T")=0
	; SENDINPUT n't
; RETURN
; ::+n`;+t::N'T
; -------------------------------------------------------------------

; OTHER WAY PAIR TO DO IT WITH CASE SENSITIVE ON
; WORKER
; -------------------------------------------------------------------
; :C:n`;t::
; :C:N`;T::
	; SENDINPUT n't
; RETURN

; OTHER WAY PAIR TO DO IT WITH CASE SENSITIVE ON
; WORKER
; -------------------------------------------------------------------
; :C:n`;t::
	; SENDINPUT n't
; RETURN
; :C:N`;T::
	; SENDINPUT n't
; RETURN
; -------------------------------------------------------------------



STRING_INVERT_MESSENGER:
	IF GetKeyState("Capslock", "T")
	{
		; TOOLTIP % "YES ON"
		 Lab_Invert_Char_Out:= ""
		 Loop % Strlen(MESSENGER_KEY) {
			Lab_Invert_Char:= Substr(MESSENGER_KEY, A_Index, 1)
			if Lab_Invert_Char is upper
			   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) + 32)
			else if Lab_Invert_Char is lower
			   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) - 32)
			else
			   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Lab_Invert_Char
		 }
		 MESSENGER_KEY=%Lab_Invert_Char_Out%
	}
	; Source Creditor
	; ----
	; Convert text - uppercase, lowercase, capitalized or inverted - Scripts and Functions - AutoHotkey Community
	; https://autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/
	; ----
RETURN


; *D::
; *::
; MSGBOX  You pressed %A_ThisHotkey%.
; Return

TIMER_HOTKEY:
	VAR_IN_NAME=Confirm Save As ahk_class #32770
	SetTitleMatchMode 3  ; Specify Full path
	IFWINEXIST %VAR_IN_NAME%
	IFWINEXIST Confirm Save As ahk_exe VB6.EXE
	{
		ControlGetText CONTROL_TEXT,Button1,%VAR_IN_NAME%
		
		STRING_V:=&&Yes  0
		IF INSTR(CONTROL_TEXT,%STRING_V%)>1
		{	
			; NA [v1.0.45+]: May improve reliability. See reliability below.
			ControlClick, Button1,%VAR_IN_NAME%,,,, NA x10 y10 
			SOUNDBEEP 4000,300
			VAR_DONE_ESCAPE_KEY=TRUE
		}
		IF CONTROL_TEXT=&Yes
		{
			Secs_MSGBOX_04=5
			SOUNDBEEP 5000,200
		}

		IF Secs_MSGBOX_04>0 	
			Secs_MSGBOX_04-=1
			
		ControlSetText,Button1,&Yes  %Secs_MSGBOX_04%, %VAR_IN_NAME%
	}
RETURN


; F4::
	; SEND ^{V}
	; SEND {ENTER}
	
	
	; ; ---------------------------------------------------------------
	; ; # Win (Windows logo key) 
	; ; ! Alt 
	; ; ^ Control 
	; ; + Shift 
	; ; & An ampersand 
	; ; ---------------------------------------------------------------
	
; RETURN

TIMER_WSCRIPT_FOCUS_LEFT_KILL:
	
	RETURN
	
	WSCRIPT_FOCUS_SET_FLAG_01=FALSE
	IfWinActive ahk_class #32770
	IfWinActive ahk_exe WScript.exe
	{	
		WSCRIPT_FOCUS_SET_FLAG_01=TRUE
		WSCRIPT_FOCUS_SET_FLAG_02=TRUE
	}
	
	IF WSCRIPT_FOCUS_SET_FLAG_01=FALSE
	IF WSCRIPT_FOCUS_SET_FLAG_02=TRUE
	{

		WinGet, List, List, ahk_exe WScript.exe
		Loop %List%  
		{ 
			Process, Close, WScript.exe
			SOUNDBEEP 1200,40
		}
	}
RETURN






; -------------------------------------------------------------------
; SESSION 00
; -------------------------------------------------------------------
; GOSUB CHECK_ESC_KEY
; -------------------------------------------------------------------
; ONLY USER UP NOT DOWN
; UP ON OWN NOT WORK -- REQUIRE A HOTKEY DOWN STATE ALSO
; BEEN THROUGH BEFORE -- *~ESC UP::
; RESULT AT LAST USE GetKeyState, state, ESC  -- UP THING A ME JIG
; NOT DOUBEL HITT
; HARD THINK ANOTHER WAY TO DO THAT THING
; SURE REPEATER KEY NOT WORK IF LOOK LAST KEY
; -------------------------------------------------------------------
; A TREAT CODER TO DO AFTER EVERY THING ELSE TODAY WAS ADD ROUTINE TO 
; ADD AN OPTION TO TOGGLE THE "SEARCH RESULTS WINDOW NOTEPAD-PLUS-PLUS
; TOOK A SEARCH TO FINDER -- CREDIT LINK BELOW
; -------------------------------------------------------------------
; Tue 17-Sep-2019 15:25:01
; Tue 17-Sep-2019 18:05:00 -- 2 HOUR
; -------------------------------------------------------------------
; *~ESC::
	; GOSUB CHECK_ESC_KEY

CHECK_ESC_KEY:
	SetTitleMatchMode 3  ; Specify Full path

	GetKeyState, state, ESC  
	if (state = "U")
		RETURN
	
	VAR_DONE_ESCAPE_KEY=FALSE

	IfWinActive ahk_class ConsoleWindowClass
	{
		WinClose, ahk_class ConsoleWindowClass
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	IfWinActive ahk_class IrfanView
	{
		WinClose, IrfanView
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	SET_GO_12=
	IfWinActive ahk_class MediaPlayerClassicW
		SET_GO_12=11
	IfWinActive ahk_class Afx:00007FF6A22C0000:b:0000000000010003:0000000000000006:0000000000000000
		SET_GO_12=11
	IfWinActive ahk_class AfxControlBar140su
		SET_GO_12=11
	IfWinActive ahk_class FullScreenClass
		SET_GO_12=11
	
	IF SET_GO_12
	{
		GOSUB KILL_ALL_MPC_EXE_NAME
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	IfWinActive Find ahk_class #32770
	IfWinActive Find ahk_exe notepad++.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	IfWinActive Replace ahk_class #32770
	IfWinActive Replace ahk_exe notepad++.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	IfWinActive Microsoft Visual Basic ahk_class #32770
	IfWinActive Microsoft Visual Basic ahk_exe vb6.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	IfWinActive Find ahk_class #32770
	IfWinActive Find ahk_exe vb6.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	; Find ahk_class #32770 ahk_exe vb6.exe
	IfWinActive Replace ahk_exe vb6.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	IfWinActive Quick Watch ahk_exe vb6.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	IfWinActive Quick Watch ahk_exe vb6.exe
	{	WinClose
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	; OK
	; Help
	; Compile error:
	; Expected: Then or GoTo
	IfWinActive Microsoft Visual Basic ahk_class #32770
	IfWinActive Microsoft Visual Basic ahk_exe VB6.EXE
	{	
		ControlGetText, Output_Var, Static2
		if instr(Output_Var,"Compile error:")
		{
			WinClose
			SoundBeep , 1500 , 400
			VAR_DONE_ESCAPE_KEY=TRUE
		}
	}

	SetTitleMatchMode 2  ; Specify Full path
	; ---------------------------------------------------------------
	; VBKeepRunner - Microsoft Visual Basic [design] - [Object Browser]
	; ---------------------------------------------------------------
	IfWinActive Microsoft Visual Basic [design] ahk_class wndclass_desked_gsk
	IfWinActive [design] - [Object Browser] ahk_class wndclass_desked_gsk
	IfWinActive Microsoft Visual Basic [design] ahk_exe VB6.EXE
	IfWinActive [design] - [Object Browser] ahk_exe VB6.EXE
	{	
		Control, Hide ,, Object Browser
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	; ---------------------------------------------------------------

	SetTitleMatchMode 3  ; Specify Full path
	IfWinActive Confirm Save As ahk_class #32770
	IfWinActive Confirm Save As ahk_exe VB6.EXE
	{	
		ControlClick, &Yes,Confirm Save As ahk_class #32770
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	; ---------------------------------------------------------------
	; IF NOT WANT A DUAL CALLER OF VB6 TO SAME CODER 
	; SOLUTION HERE -- 01 OF 02 -- AND -- 02 OF 02 -- NEXT
	; ---------------------------------------------------------------
	IfWinActive New Project ahk_class #32770
	IfWinActive New Project ahk_exe VB6.EXE
	{	
		WinClose
		SoundBeep , 5000 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}
	
	; 02 OF 02
	; ---------------------------------------------------------------
	IfWinActive Microsoft Visual Basic ahk_class wndclass_desked_gsk
	IfWinActive Microsoft Visual Basic ahk_exe VB6.EXE
	; DO BOTH TOGETHER OR ONE AT A GO
	; IF VAR_DONE_ESCAPE_KEY=FALSE
	{	
		SEND !{F4}
		SoundBeep , 6000 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	SetTitleMatchMode 2
	IfWinActive ] Options ahk_class #32770
	IfWinActive ] Options ahk_exe GoodSync-v10.exe
	{	
		ControlGetPos, x, y, , , Button63, Options ahk_class #32770 ; SAVE BUTTON
		MouseMove, X+10, Y+10
		ControlClick, Button63, Options ahk_class #32770,,,, NA x20 y20
		SoundBeep , 1500 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	IfWinActive Rename Job ahk_class #32770
	IfWinActive Rename Job ahk_exe GoodSync-v10.exe
	{	
		WinClose
		SoundBeep , 5000 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	IfWinActive Left Folder ahk_class #32770
	IfWinActive Left Folder ahk_exe GoodSync2Go.exe
	{	
		WinClose
		SoundBeep , 5000 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	
	IfWinActive ahk_class #32770
	IfWinActive ahk_exe WScript.exe
	{	
		; WinGet, List, List, ahk_class AutoHotkey 
		WinGet, List, List, ahk_exe WScript.exe
		Loop %List%  
		{ 
			Process, Close, WScript.exe
			SOUNDBEEP 1200,40
		}

		; PROCESS, CLOSE, WScript.exe
		; SoundBeep , 5000 , 400
		VAR_DONE_ESCAPE_KEY=TRUE
	}

	
	
	; ---------------------------------------------------------------
	; WINAMP VISUALIZATION CONTROL KEY WINDOW
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	SetTitleMatchMode 3
	GetKeyState, state, Shift
    if state = U
	{
		SET_GO=FALSE
		IfWinActive ahk_class Winamp Gen
		{	
			WinActivate ahk_class Winamp v1.x
			SET_GO=TRUE
		}
		IfWinActive ahk_class Winamp v1.x
			SET_GO=TRUE
		IfWinActive ahk_class Winamp PE
			SET_GO=TRUE
		IF SET_GO=TRUE
		IfWinActive ahk_exe winamp.exe
		{	
			; SEND {Ctrl down}{Shift down}K{Shift up}{Ctrl up}
			SEND ^+K
			SoundBeep , 9000 , 100
			VAR_DONE_ESCAPE_KEY=TRUE
		}
	}
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	GetKeyState, state, Shift
	if state = D
	IfWinActive ahk_class Winamp Gen
	{
		SoundBeep , 1000 , 100
		WinGet, HWND_10, ID, ahk_class Winamp Gen
		VAR_DONE_01=FALSE
		if isWindowFullScreen(%HWND_10%)=0 
		{
			SEND !{ENTER}
			VAR_DONE_01=TRUE
			VAR_DONE_ESCAPE_KEY=TRUE
		}
		if isWindowFullScreen(%HWND_10%)<>0 
		IF VAR_DONE_01=FALSE
		{
			SEND !{D}
			VAR_DONE_ESCAPE_KEY=TRUE
		}
	}
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; WINAMP VISUALIZATION CONTROL KEY WINDOW
	; ---------------------------------------------------------------

	ESCAPE_KEY_COUNT+=1
 
	VAR_DONE_ESCAPE_NOTEPAD=
	; ---------------------------------------------------------------
	DetectHiddenWindows, OFF
	SetTitleMatchMode 2
	IfWinActive ahk_class Notepad++
	IfWinActive ahk_exe notepad++.exe
	{	
		; -------------------------------------------------------
		; IF THE RESULTS PANE IS OPEN, CLOSE IT
		; BUTTON1 IS THE CLASS NAME FOR THE TITLE BAR AND 
		; CLOSE BUTTON OF THE RESULTS PANE WHEN DOCKED
		; -------------------------------------------------------
		ControlGet, OutputVar, Visible,, Button1, Notepad++
		if ErrorLevel = 0
			If OutputVar=1
			{
				; FOUND IT DOCKED
				; GET THE SIZE AND COORDINATES OF THE TITLE BAR AND BUTTON
				ControlGetPos, X, Y, Width, Height, Button1
				; SET THE COORDINATES OF THE CLOSE BUTTON
				X := WIDTH - 9
				Y := 5
				; SEND A CLICK
				ControlClick, Button1,,,,, NA x%X% y%Y%
				SoundBeep , 4000 , 50
				VAR_DONE_ESCAPE_KEY=TRUE
				VAR_DONE_ESCAPE_NOTEPAD=TRUE
				LOOP, 10000
				{
					ControlGet, OutputVar, Visible,, Button1, Notepad++
					If OutputVar=0
						BREAK
					SLEEP 100
				}
				SLEEP 100
			}
			IF !VAR_DONE_ESCAPE_NOTEPAD
			{
				ControlGet, OutputVar, Visible,, Button1, Notepad++
				If OutputVar=0
				{
					; ----------------------------------------------------------------------------------
					; FROM HERE ONLY WANT CLOSE NOT AS F7 KEY OPEN AGAIN -- THE FLIP FLOP IS TOO MUCH
					; IDEAL IF SEARCH RESULT FIND SOMETHING TO ISSUE FIND RESULT ITEM
					; ALLOW ESCAPE CLOSE AND FOR TIMER ALLOW QUICK REOPEN ESC KEY RATHER THAN F7 AS HERE
					; OPEN SEARCH RESULT AGAIN WOULD REQUIRE SEARCH AGAIN
					; CLOSE SEARCH BOX ISSUE ESCAPE OPEN RESULT BOTH MAYBE CHECK THERE
					; Sat 30-May-2020 16:51:00
					; ----------------------------------------------------------------------------------
					; SendInput {F7}
					; SoundBeep , 2000 , 100
					VAR_DONE_ESCAPE_KEY=TRUE
				}
				LOOP, 10000
				{
					ControlGet, OutputVar, Visible,, Button1, Notepad++
					If OutputVar=1
						BREAK
					SLEEP 100
				}
				SLEEP 100
			}
			; ---------------------------------------------------
			; ESCAPE KEY FOR RUN AWAY
			; ---------------------------------------------------
	}
	

	; ---------------------------------------------------------------
	; 003
	; ----
	; CREATE A HOTKEY /KEYBOARD SHORTCUT TO CLOSE THE NOTEPAD++ FIND RESULTS WINDOW - SUPER USER
	; https://superuser.com/questions/700357/create-a-hotkey-keyboard-shortcut-to-close-the-notepad-find-results-window
	; ----
	; ---------------------------------------------------------------
	; 002
	; ----
	; ADD AN OPTION TO TOGGLE THE "SEARCH RESULTS WINDOW" · ISSUE #2946 · NOTEPAD-PLUS-PLUS/NOTEPAD-PLUS-PLUS · GITHUB
	; https://github.com/notepad-plus-plus/notepad-plus-plus/issues/2946
	; ----
	; ---------------------------------------------------------------
	; 001
	; ----
	; AHK CLOSE NOTEPAD++ SEARC BOX - GOOGLE SEARCH
	; https://www.google.com/search?q=AHK+CLOSE+NOTEPAD%2B%2B+SEARC+BOX&oq=AHK+CLOSE+NOTEPAD%2B%2B+SEARC+BOX&aqs=chrome..69i57.10623j1j7&sourceid=chrome&ie=UTF-8
	; ----

	
	; ; 
	; ; ---------------------------------------------------------------
	; ; ---------------------------------------------------------------
	; ; 
	; ; THIS IS MY CODER TO BRING UP WINDOW 
	; ; VB_KEEP_RUNNER 
	; ; AND
	; ; ELITE SPY 
	; ; BUT FOR NOW THESE TWO PROGRAM WHEN COMPUTER UNDER PRESSURE FOR LONG WHEN GET BACK TO
	; ; THEY BOTH NEVER SEEM RUNNER
	; ; SO SOMEHOW GOT TO FIND THAT WINDOW POP UP NOT HAPPEN AS NORMAL
	; ; AND THEN TAKE ACTION TO KILL AND RELOAD
	; ; I TRY AND CATCH ERROR OF WINDOW-STATE STYLE
	; ; AND RESULT WAS 0
	; ; 0 FOR NONE OF MIN OR MAX AS IF WAS NORMAL
	; ; BUT NOT TRUE 0 WAS RESULT BUT NOT SHOW-ER
	; ; SO NEXT TIME DEBUG CHECK FOR HIDDEN STATE OR MORE
	; ; FROM       Sat 11-May-2019 06:04:11
	; ; STOP TIME  Sat 11-May-2019 07:07:55 -- CODE SOMETHING ELSE -- VB_KEEP_RUNNER -- IMPROVE
	; ; RESUME-A   Sat 11-May-2019 08:24:53 --
	; ; TO         Sat 11-May-2019 09:41:33
	; ; ---------------------------------------------------------------
	; ; ---------------------------------------------------------------
	; VB_KEEP_RUNNER_VAR=FALSE
	; VB_KEEP_RUNNER_TITLE=VB_KEEP_RUNNER
	; GetKeyState, state, Shift
	; if state = D
	; IfWinExist, %VB_KEEP_RUNNER_TITLE%
	; {
		; WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		; WinGet style, MinMax, ahk_id %HWND_10%
		; IF style=-1
		; {
			; WinRestore, %VB_KEEP_RUNNER_TITLE%
			; WinActivate, %VB_KEEP_RUNNER_TITLE%
			; SoundBeep , 1000 , 100
			; SoundBeep , 2000 , 100
			; SoundBeep , 3000 , 100
			; VAR_DONE_ESCAPE_KEY=TRUE
			; SLEEP 400
		; }
		; WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		; WinGet style, MinMax, ahk_id %HWND_10%
		; IF style=0
		; {
			; VAR_DONE_ESCAPE_KEY=TRUE
			; ; VB_KEEP_RUNNER_VAR=TRUE
			; ; VB_KEEP_RUNNER_VAR_2=TRUE
			; ; SLEEP 400
		; }
	; }
	; ; ---------------------------------------------------------------
	; ; ---------------------------------------------------------------
	; SetTitleMatchMode 2
	; VB_ELITE_SPY_VAR=FALSE
	; ELITE_SPY_TITLE=EliteSpy`+ 2001 __ www.PlanetSourceCode.com __ Version
	; GetKeyState, state, Shift
	; if state = D
	; IfWinExist, %ELITE_SPY_TITLE%
	; {
		; WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		; WinGet style, MinMax, ahk_id %HWND_10%
		; IF style=-1
		; {
			; WinRestore, %ELITE_SPY_TITLE%
	 		; WinActivate, %ELITE_SPY_TITLE%
			; SoundBeep , 1000 , 100
			; SoundBeep , 2000 , 100
			; SoundBeep , 3000 , 100
			; VAR_DONE_ESCAPE_KEY=TRUE
			; VB_ELITE_SPY_VAR=TRUE
		; }
		; WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		; WinGet style, MinMax, ahk_id %HWND_10%
		; IF style=0
		; {
			; VAR_DONE_ESCAPE_KEY=TRUE
			; ; VB_ELITE_SPY_VAR=TRUE
			; ; VB_ELITE_SPY_VAR_2=TRUE
			; ; SLEEP 400
		; }
	; }
	
	; VB_KEEP_RUNNER_VAR=FALSE
	; VB_KEEP_RUNNER_VAR_2=FALSE
	; IF VB_KEEP_RUNNER_VAR=TRUE
	; {
		; LOOP 
		; {
			; X_COUNTER+=1
			; WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
			; WinGet style, MinMax, ahk_id %HWND_10%
			; ; IF style=0
			; ; MSGBOX % style
			; ; MinMax
			; ; Retrieves the minimized/maximized state for a window.
			; ; WinGet, OutputVar, MinMax [, WinTitle, WinText, ExcludeTitle, ExcludeText]
			; ; OutputVar is made blank if no matching window exists; otherwise, it is set to one of the following numbers:
			; ; •-1: The window is minimized (WinRestore can unminimize it).
			; ; •1: The window is maximized (WinRestore can unmaximize it).
			; ; •0: The window is neither minimized nor maximized.


			; IF style=-1
			; {
				; VB_KEEP_RUNNER_VAR_2=TRUE
				; BREAK
			; }
			; SLEEP 100
			; IF X_COUNTER>100
				; BREAK
		; }
	; }
	; IfWinNotExist %VB_KEEP_RUNNER_TITLE%
	; {
		; VAR_DONE_ESCAPE_KEY=TRUE
		; SoundBeep , 3000 , 100
		; FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
		; IfExist, %FN_VAR%
			; {
				; Run, %FN_VAR% MAXIMUM
			; }
	; }	
	
	; IF VB_KEEP_RUNNER_VAR_2=TRUE
	; {
		; VAR_DONE_ESCAPE_KEY=TRUE
		; WinGet, HWND_10, ID, %VB_KEEP_RUNNER_TITLE%
		; WinGet, UniquePID, PID,  ahk_id %HWND_10%
		; IF UniquePID>0
		; {
			; SOUNDBEEP 1000,100
			; Process, Close, %UniquePID% 
		; }

		; IfWinNotExist %VB_KEEP_RUNNER_TITLE%
		; {
			; SoundBeep , 3000 , 100
			; FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
			; IfExist, %FN_VAR%
				; {
					; Run, %FN_VAR% MAXIMUM
				; }
		; }	
	; }
	
	; VB_ELITE_SPY_VAR=FALSE
	; VB_ELITE_SPY_VAR_2=FALSE
	; IF VB_ELITE_SPY_VAR=TRUE
	; {
		; LOOP 
		; {
			; X_COUNTER+=1
			; WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
			; WinGet style, MinMax, ahk_id %HWND_10%
			; ; IF style=0
			; ; MSGBOX % style

			; IF style=0
			; {
				; VB_ELITE_SPY_VAR_2=TRUE
				; BREAK
			; }
			; SLEEP 100
			; IF X_COUNTER>100
				; BREAK
		; }
	; }
	; IfWinNotExist %ELITE_SPY_TITLE%
	; {
		; VAR_DONE_ESCAPE_KEY=TRUE
		; SoundBeep , 3000 , 100
		; FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
		; IfExist, %FN_VAR%
			; {
				; Run, %FN_VAR% MAXIMUM
			; }
	; }	
	
	; IF VB_ELITE_SPY_VAR_2=TRUE
	; {
		; VAR_DONE_ESCAPE_KEY=TRUE
		; WinGet, HWND_10, ID, %ELITE_SPY_TITLE%
		; WinGet, UniquePID, PID,  ahk_id %HWND_10%
		; IF UniquePID>0
		; {
			; SOUNDBEEP 1000,100
			; Process, Close, %UniquePID% 
		; }

		; IfWinNotExist %ELITE_SPY_TITLE%
		; {
			; SoundBeep , 3000 , 100
			; FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
			; IfExist, %FN_VAR%
				; {
					; Run, %FN_VAR% MAXIMUM
				; }
		; }	
	; }
	
	; ; ---------------------------------------------------------------
	; ; # Win (Windows logo key) 
	; ; ! Alt 
	; ; ^ Control 
	; ; + Shift 
	; ; & An ampersand 
	; ; ---------------------------------------------------------------
	
	IF VAR_DONE_ESCAPE_KEY=FALSE
	{
		SOUNDBEEP 4000,50
		SOUNDBEEP 5000,40
	}
	
RETURN




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


; -------------------------------------------------------------------
; REPLACE F10 TO DO CONTROL PRINT SCREEN
; FOR CLIPBOARD SCREEN SHOT -- 
; CODE PICPICK WON'T HOT KEY F10 ON WIN 7 AND UP
; PIKPICK SETTER -- CAPTURE FULL SCREEN IS NONE AND CARRY ON LIKE NORMAL CTRL PRNT SCRN
; PIKPICK SETTER -- CAPTURE WINDOW FORM-ER AREA IS SHIFT F10 WITH THIS CODE 
; LINE ABOVE THE NORM IS PRESS F10 THIS CODE TRANSLATE TO SHIFTER F10
; SHIFT IS +
; +F10 IS ACTIVE WINDOW 
; CTRL PRINTSCR IS WHOLE SCREEN OF PIKPICKER
; ALL REST SET TO NONE HOT KEY
; ^ CTRL IS BETTER THAN SHIFT DRAW A DIALOG BOX UP SOME INSTACE PDF ACROBAT
; set pikpick to ctrl f10 area and ctrl prt scrn for normAl same as before
; Sendinput +{PrintScreen} -- NOT USER
; ^ CTRL F10 PRT SCRN
; -------------------------------------------------------------------
;F10::^F10

F10::
	SoundBeep , 1500 , 400
	;SetCapsLockState , Off
	;SetNumLockState , Off
	;SetScrollLockState , Off

	Sendinput ^{F10}
	; LOGITECH MOUSE INFO POP UP WONT LIKE CAPS CHANGE OR HOT KEY F10 THINK INFO REQUIRE ABOUT CAPITAL CHANGE OR TURN LOGITECH NOTIFY OFF
Return
; -------------------------------------------------------------------
; -------------------------------------------------------------------


TIMER_ENTER:
	DetectHiddenWindows, oFF
	SetTitleMatchMode 3  ; Specify Full path.
	IfWinExist, Clone Job ahk_class #32770
	{
		ControlClick, OK, Clone Job ahk_class #32770
		SoundBeep , 2500 , 100
	}
RETURN


; *C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk - Notepad++ [Administrator]

; WINGET Autokey -- 41-Minimize Chrome Close & Close RButton.ahk

; -------------------------------------------------------------------
; ADD THE TERMINATOR VERSION NUMBER AND THEN WE ARE ABLE TO USE EXACT 
; STRING MATCHING IN CASE NOTEPAD HAD IT
; YOU HEARD IT HEAR FIRST
; -------------------------------------------------------------------

; F5::^F5


; F12::
; {
	; SetTitleMatchMode 3  ; EXACTLY
	; DetectHiddenWindows, ON
	; AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
	; AHK_TERMINATOR_VERSION:=StrReplace(AHK_TERMINATOR_VERSION, """" , "")
	; FILE:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk"
	; WinGet, UniquePID, PID, %FILE%%AHK_TERMINATOR_VERSION%
	
	; IF UniquePID>0
	; {
		; SOUNDBEEP 1000,100
		; ; WinKill, Ahk_PID %UniquePID% 
		; Process, Close, %UniquePID% 
	; }
	; FILE:="Autokey -- 41-Minimize Chrome Close & Close RButton.ahk"
	; WinGet, UniquePID, PID, %FILE%
	
	; IF UniquePID>0
	; {
		; SOUNDBEEP 1000,100
		; ; WinKill, Ahk_PID %UniquePID% 
		; Process, Close, %UniquePID% 
	; }
; }
; RETURN

; F12::Process, Close,WScript.exe
; C:\Windows\SysWOW64\WScript.exe


; -------------------------------------------------------------------
; THE FACEBOOK ONE IS HERE NOW
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 40-Auto Add Photo Name Messenger Facebook.ahk
; FACE LIKE A BOOK
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; HERE FOR FACEBOOK 
; GOT SOME PHOTO IN FOLDER
; THIS WILL GET FILENAME WITH ANY DESCRIPTION
; AND PUT IT IN FACEBOOK 
; SO YOU DON'T HAVE TO TYPE ANY IN ADD YOUR DESCRIPTION ON COMPUTER 
; FIRST AND DON'T HAVE TO MAKE EXTRA COPY
; WISH SPACE BOOK GRIN BOOK _ AI _ WOULD DO SOMETHING LIKE THAT
; THOUGH GUESS THEY ARE SECRETLY KEEPING ALL YOUR FILENAME INFO WITH EACH PHOTO SO IT DON'T GET LOST OR SOMETHING
; -------------------------------------------------------------------
; REM IN AND OUT WHEN WANT TO USE IT 
; SUPPLY PATH NAME _ UP A BIT
; MIGHT MAKE IT OWN AHK FILE THIS ONE AS USED A LOT AND ONLY WANT 
; CALL UP WHEN USER
; MY FAV HOTKEY IS F4 F5 AND OTHER BUT HARDLY USE MANY OR I GET IT TOO COMPLICATED
; GUESSING WHERE MY FAVOURITE HOTKEY ARE
; USE THEM WHEN YOU WANT
; WRITTEN WELL BEFORE THIS DATE
; JUST USE AGAIN TODAY
; -------------------------------------------------------------------
; [ Tuesday 21:31:30 Pm_09 October 2018 ]
; -------------------------------------------------------------------


; CPC
; XXXF4::
; F4::
{
	TITLE_ADD=
	SEXY=0
	; ----
	; Activate a chrome tab by its name - Ask for Help - AutoHotkey Community
	; https://autohotkey.com/board/topic/110607-activate-a-chrome-tab-by-its-name/
	; ----
	
	WinActivate, ahk_class Chrome_WidgetWin_1
	WinWaitActive, ahk_class Chrome_WidgetWin_1
	WinGetTitle, CurrentWindowTitle, ahk_class Chrome_WidgetWin_1

	SET_GO=FALSE
	If CurrentWindowTitle contains 404 Page Not Found
		SET_GO=TRUE
	If CurrentWindowTitle contains Raspberry Pi - Google Chrome
		SET_GO=TRUE
		
	IF SET_GO=FALSE
		RETURN
				
	Loop
	{
		; TAB BEFORE
		; --------
		Send, ^{Tab}    ; ---- CONTROL TAB
		Sleep, 200
		WinActivate, ahk_class Chrome_WidgetWin_1
		IfWinNotExist, ahk_class Chrome_WidgetWin_1
			Return
		WinGetTitle, CurrentWindowTitle, ahk_class Chrome_WidgetWin_1
		SET_GO=FALSE
		If CurrentWindowTitle contains 404 Page Not Found
			SET_GO=TRUE
		If CurrentWindowTitle contains CPC - Google Chrome
			SET_GO=TRUE
		If CurrentWindowTitle contains New Tab - Google Chrome
			SET_GO=TRUE
			
		; CPC UK | Electronic Parts & Components, Raspberry Pi - Google Chrome
		If CurrentWindowTitle contains Raspberry Pi - Google Chrome
			SET_GO=TRUE
			
		IF SET_GO=TRUE
		{
			Sleep, 200
			; KILL PAGE
			; ---------
			SendInput ^w
			SEXY=0
		}
		
		If TITLE_ADD not contains %CurrentWindowTitle%
			SEXY=0

		SEXY+=1

		IF SEXY>30
		{
			SOUNDBEEP 3000,100
			SOUNDBEEP 2000,100
			BREAK
		}
			
		TITLE_ADD=%TITLE_ADD%%CurrentWindowTitle%
	}
}
Return


; ^/::
; {
	; SOUNDBEEP 1000,50
	; SENDINPUT, {F5}
; }
; RETURN


; C:\SCRIPTER\NOTEPAD TALK\TEXT 2018-11-25 __ D DRIVE FOLDER NAME.txt

F5_ROUTINE:

	; SET TO 0 FOR 1ST BEGINNER
	; HELP KEEP IN SYNC ADJUST AS GO
	VAR_COUNTER=0
	; -------------------------------------------------------------------
	; -------------------------------------------------------------------

	; GLOBAL FILE_SCRIPT
	; GLOBAL FILE_SCRIPT_COUNT

	FILE_SCRIPT_COUNT=0
	FILE_SCRIPT := Object()
	FOLDER_EXCLUDE := Object()

	; FILE_SCRIPT := []

	; -------------------------------------------------------------------
	; SET THE PATH OF PHOTO FOLDER _ WITHOUT RECURSING SUB-FOLDER IS GOOD IDEA
	; TO BE ABLE QUICKLY COUNT THEM VERIFY AND SEE AS ORDER TO GO ONTO FACEBOOK
	; AND NOT STRETCH MY CODE TOO MUCH ABOUT WANT RECURSING SUB-FOLDER 
	; SINGLE FOLDER ONLY AT THE MOMENT
	; -------------------------------------------------------------------
	
	; Example #3: Retrieve file names sorted by name (see next example to sort by date):
	FileList =  ; Initialize to be blank.
	Loop, Files, M:\*.*, D
		FileList = %FileList%%A_LoopFileName%`n
	Sort, FileList ;  R  ; The R option sorts in reverse order. See Sort for other options.

	Loop, parse, FileList, `n
	{
		if A_LoopField =  ; Ignore the blank item at the end of the list.
			continue
		
		FILE_SCRIPT[A_Index] := A_LoopField
		FILE_SCRIPT_COUNT := A_Index
		; MSGBOX % FILE_SCRIPT[A_Index]
		
	}	

	Loop % FILE_SCRIPT.MaxIndex()
	{
		FILE_NAME:=FILE_SCRIPT[A_Index]
		StringUpper, FILE_NAME,FILE_NAME
		FILE_SCRIPT[A_Index]:=FILE_NAME
	}

	
	ACUM=0
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "$RECYCLE.BIN"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "_gsdata_"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "System Volume Information"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "$GetCurrent"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "AWS"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "Documents and Settings"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "Program Files (x86)"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "ProgramData"
    ; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "RECOVERY"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "SADPLOG"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "windows"
	
	
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "# MY DOCS"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "#0 1 INSTALLATIONS"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 ART"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 ART LOGGERS"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 ART LOGGERS - WEBCAM"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 HTTrack"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 MUSIC ---"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 MOBILE-1"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "0 00 MOBILE-2"
	; ; ACUM+=1
	; ; FOLDER_EXCLUDE[ACUM] := "0 00 VIDEO SNAPSHOT CCSE HIKVISION"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "03_MICROSOFT"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "BT Cloud Sync OLD"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "DL"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "DSC"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "KAT MP3 RECORDER"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "VB6"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "VB6-ARC"
	; ; ACUM+=1
	; ; FOLDER_EXCLUDE[ACUM] := "VB6 _ IN SYNC"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "VB6-EXE"
	
	
	Loop % FOLDER_EXCLUDE.MaxIndex()
	{
		FILE_NAME:=FOLDER_EXCLUDE[A_Index]
		StringUpper, FILE_NAME,FILE_NAME
		FOLDER_EXCLUDE[A_Index]:=FILE_NAME
	}
	
	
	X_Index:=FILE_SCRIPT.MaxIndex()
	Loop % FILE_SCRIPT.MaxIndex()
	{
		X_Index-=1
		Loop % FOLDER_EXCLUDE.MaxIndex()
		{
			IF (InStr(FILE_SCRIPT[X_Index]"*", FOLDER_EXCLUDE[A_Index]"*"))
			{
				; msgbox % FILE_SCRIPT[X_Index]
				FILE_SCRIPT.RemoveAt(X_Index)
			}
		}	
	}	

	; --------------------------------------------------
	; REQUIRE OPTION TO IGNORE IF FOLDER OLDER THAN DATE
	; LIKE ABOVE 1 ROUTINE
	; --------------------------------------------------

	
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	
	; 1ST
	; SETTIMER AUTO_CLONE_JOB, 1000

	; 2ND _ GOT TO SET THE PATH FOR EACH JOB
	; SETTIMER SET_OK_BOX,100
	
	; 3RD
	; GOSUB DISPLAY_TOOLTIP
	; SETTIMER RENAME_JOBS_FROM_DIRECTORY_SCANNER,1000
	
	; 4TH
	; SETTIMER RENAME_PATH_OF_JOBS_LEFT_OR_RIGHT,100
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------

RETURN


; -------------------------------------------------------------------
; -------------------------------------------------------------------
STATE_XYPOS_LIMIT:
	STATE_XYPOS_COUNTER=0
	OLD_STATE_XYPOS_COUNTER=0
	SETTIMER, STATE_XYPOS_LIMIT,OFF
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
CHECK_NEW_WINDOW_TIMER_TOP_LEFT_MOUSE_CLOSE_MPC_SOUNDPLAY:
	SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
	SETTIMER CHECK_NEW_WINDOW_TIMER_TOP_LEFT_MOUSE_CLOSE_MPC_SOUNDPLAY,OFF
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
TIMER_TOP_LEFT_MOUSE_CLOSE_MPC:
{
	CoordMode, Mouse, Screen
	MouseGetPos, xpos, ypos 
	
	STATE_XYPOS:=xpos+ypos
	
	IF (STATE_XYPOS=0 and OLD_STATE_XYPOS<>%STATE_XYPOS%)
		STATE_XYPOS_COUNTER+=1
	
	OLD_STATE_XYPOS=%STATE_XYPOS%
	
	SET_GO=FALSE
	
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_NAME, A

	XR=0
	IF INSTR(CLASS,"MediaPlayerClassicW")
		XR=1
	IF INSTR(CLASS,"Afx:00007FF6A22C0000:b:0000000000010003:0000000000000006:0000000000000000")
		XR=2
	IF INSTR(CLASS,"AfxControlBar140su")
		XR=3
	IF INSTR(CLASS,"FullScreenClass")
		XR=4
	IF INSTR(CLASS,"IrfanView")
		XR=5
	
	IF XR>0
	IF STATE_XYPOS=0
	IF STATE_XYPOS_COUNTER<>%OLD_STATE_XYPOS_COUNTER%
		SOUNDBEEP 2000,100
		
	IF STATE_XYPOS_COUNTER>0
	IF STATE_XYPOS_COUNTER<>%OLD_STATE_XYPOS_COUNTER%
	{
		OLD_STATE_XYPOS_COUNTER=%STATE_XYPOS_COUNTER%
		SetTimer, STATE_XYPOS_Limit, 2000 ;<-- set a oneshot 1 second timer to stop the loop
	}
	OLD_STATE_XYPOS_COUNTER=%STATE_XYPOS_COUNTER%
	
	IF (STATE_XYPOS_COUNTER>1 or A_PriorKey=27)
	{
		IF XR>0 
		{
			; TOOLTIP % HWND_X "," HWND_ACTIVE "," SET_GO
			IF XR<4
			{
				SETTIMER CHECK_NEW_WINDOW_TIMER_TOP_LEFT_MOUSE_CLOSE_MPC_SOUNDPLAY,50
				GOSUB KILL_ALL_MPC_EXE_NAME
				XR=0
				STATE_XYPOS_COUNTER=0
			}
			IF XR>0
			{
				; Process, Close, i_view32.exe
				; ---------------------------------------------------
				; ABOVE METHOD FAIL DESTROYS WINDOWS EXPLORER TASKBAR 
				; HAVE TO RESET EXPLORER AGAIN
				; ---------------------------------------------------

				; PostMessage, 0x112, 0xF060,,, TITLE_NAME, WinText  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
				; PostMessage, 0x112, 0xF060,,, TITLE_NAME , CLASS  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
				; PostMessage, 0x112, 0xF060,,, TITLE_NAME ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
				
				; ---------------------------------------------------
				; BOTH METHOD WORK BUT HARD TO USE WITH PASS VARIABLE 
				; UNLESS DEBUG A BIT
				; ---------------------------------------------------
				SETTIMER CHECK_NEW_WINDOW_TIMER_TOP_LEFT_MOUSE_CLOSE_MPC_SOUNDPLAY,50
				PostMessage, 0x112, 0xF060,,, IrfanView ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
				WinClose, IrfanView
				STATE_XYPOS_COUNTER=0
			}
		}	
	}	
}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
STATE_XYPOS_LIMIT_TIMER_TOP_LEFT_MOUSE_REFRESH_BROWSER:
	STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH=0
	OLD_STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH=0
	SETTIMER, STATE_XYPOS_LIMIT_TIMER_TOP_LEFT_MOUSE_REFRESH_BROWSER,OFF
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
CHECK_NEW_WINDOW_TIMER_TOP_MOUSE_REFRESH_BROWSER_SOUNDPLAY:
	SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
	SETTIMER CHECK_NEW_WINDOW_TIMER_TOP_MOUSE_REFRESH_BROWSER_SOUNDPLAY,OFF
RETURN
CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__4G:
	SET_GO_AR=
	IF A_ComputerName=4-ASUS-GL522VW
		SET_GO_AR=TRUE
	IF A_ComputerName=7-ASUS-GL522VW
		SET_GO_AR=TRUE
	IF A_ComputerName=8-MSI-GP62M-7RD
		SET_GO_AR=TRUE
	
	IF SET_GO_AR
	{
		;SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		SENDINPUT {ESC}
	}
	SETTIMER CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__4G,OFF
RETURN
CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__1X:
	IF A_ComputerName=1-ASUS-X5DIJ
	{
		; SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		SENDINPUT {ESC}
	}
	SETTIMER CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__1X,OFF
RETURN

; -------------------------------------------------------------------
; -------------------------------------------------------------------
TIMER_TOP_LEFT_MOUSE_REFRESH_BROWSER:
{
	CoordMode, Mouse, Screen
	MouseGetPos, xpos, ypos 
	
	STATE_XYPOS_TOP_MOUSE_REFRESH:=ypos
	
	IF (STATE_XYPOS_TOP_MOUSE_REFRESH=0 and OLD_STATE_XYPOS_TOP_MOUSE_REFRESH<>%STATE_XYPOS_TOP_MOUSE_REFRESH%)
	{
		STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH+=1
		SETTIMER STATE_XYPOS_LIMIT_TIMER_TOP_LEFT_MOUSE_REFRESH_BROWSER,1000  ; PUT MOUSE UP TOP AGAIN IN 500 MILLISECOND
	}
	
	OLD_STATE_XYPOS_TOP_MOUSE_REFRESH=%STATE_XYPOS_TOP_MOUSE_REFRESH%
	
	WinGetCLASS, CLASS, A
	WinGetTITLE, TITLE_NAME, A

	XR=0
	IF INSTR(CLASS,"Chrome_WidgetWin_1")
	IF INSTR(TITLE_NAME,"Your notifications - Google Chrome")
		XR=1
	
	IF (STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH>1 or A_PriorKey=27)
	{
		IF XR>0 
		{
			SETTIMER CHECK_NEW_WINDOW_TIMER_TOP_MOUSE_REFRESH_BROWSER_SOUNDPLAY,1
			SETTIMER CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__4G,10000
			SETTIMER CHECK_NEW_WINDOW_TIMER_BROWSER_PAGE_LOAD_THEN_ESCAPE_MSGBOX_ABOUT_SPAM_DOS__1X,10000
			SENDINPUT {F5}
			STATE_XYPOS_COUNTER_TOP_MOUSE_REFRESH=0
		}	
	}	
}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------







KILL_ALL_MPC_EXE_NAME:

	; ---------------------------------------------------------------
	; WORKER AND ROUTINE THAT CALL BOTH THEM
	; ESC KEY AND MOUSE LEFT TO CORNER COUPLE TIME TO TRIG
	; NOW DO KILL ALL MPC PROCESS
	; ---------------------------------------------------------------
	; 
	; Thu 30-Jan-2020 23:03:29
	; Thu 30-Jan-2020 23:33:30 -- 30 MINUTE
	; ---------------------------------------------------------------
	
	SOUNDBEEP 1000,100
	WINGET, LIST, LIST, AHK_EXE mpc-hc64.exe
	LOOP %LIST%
	{ 
		WINGET, PID_8, PID, % "AHK_ID " LIST%A_INDEX%
		IF PID_8
			PROCESS, CLOSE, %PID_8%

		SOUNDBEEP 1200,40
	}
RETURN
; -------------------------------------------------------------------
	

; -------------------------------------------------------------------
; SET OKAY BOX AFTER MADE SELECTION
SET_OK_BOX:
{

	WinGet, HWND_1, ID, GoodSync ahk_class #32770
	IF HWND_1
	{
		WinGetText OutputVar_3,ahk_id %HWND_1%

		WinGet, HWND_2, ID, A
		IF HWND_2<>%HWND_1%
		{
			#WinActivateForce, ahk_id %HWND_1%
		}
		
		IfInString, OutputVar_3, Removable drive with volume name
		{
			LOOP
			{
				SLEEP 50
				ControlGetPos, x, y, w, h, Yes, ahk_id %HWND_1%
				MouseMove, X+10, Y+10
				
				ControlClick, Yes,ahk_id %HWND_1%
				IfWinNotExist, ahk_id %HWND_1%
					BREAK
				SoundBeep , 4000 , 50
			}
		}
	}

	WinGet, HWND_1, ID, Left Folder ahk_class #32770
	IF !HWND_1 
		WinGet, HWND_1, ID, Right Folder ahk_class #32770
	IF !HWND_1 
		O_OutputVar_store=
	
	IF HWND_1 
	{
	WinGetPos, X, Y, Width, Height, ahk_id %HWND_1%
	; tooltip % x " -- " y
	IF X<>770 and Y<>180
		WinMove, ahk_id %HWND_1%,,990,180

	ControlGetText, OutputVar_store, Edit1, ahk_id %HWND_1%
	
	; TOOLTIP % "1 " OutputVar_store "`n2 " O_OutputVar_store
	
	if O_OutputVar_store
		if OutputVar_store<>%O_OutputVar_store%
		{
			O_OutputVar_store=
			LOOP
			{
				SLEEP 50
				ControlGetPos, x, y, w, h, OK, ahk_id %HWND_1%
				MouseMove, X+10, Y+10
				
				ControlClick, OK, ahk_id %HWND_1%
				IfWinNotExist, ahk_id %HWND_1%
					BREAK
				SoundBeep , 4000 , 50
			}
		}
	; TOOLTIP % "1 " OutputVar_store "`n2 " O_OutputVar_store
	
	O_OutputVar_store=%OutputVar_store%

	}
	
}
RETURN


DISPLAY_TOOLTIP:
	VAR_COUNTER2:=VAR_COUNTER
	VAR_COUNTER2+=1
	PART_RENAME_VAR:="HDD CLOUD GD2TB "
	TOOLTIP % VAR_COUNTER2 " -- " PART_RENAME_VAR FILE_SCRIPT[VAR_COUNTER2]"`n"FILE_SCRIPT[VAR_COUNTER2],1300,50
RETURN


;; *F2::
	IF !xpos
	MouseGetPos, xpos, ypos 
	SEND {F2}
RETURN

RENAME_JOBS_FROM_DIRECTORY_SCANNER:
{

	WinGet, HWND_1, ID, Rename ahk_class #32770

	IF !HWND_1
	{
		O_ID=0
		RETURN
	}
	
	IF HWND_1=%O_ID%
		RETURN
	
	O_ID:=HWND_1
	
	VAR_COUNTER+=1
	FILE_NAME := % FILE_SCRIPT[VAR_COUNTER]
	F_2:=""
	FILE_NAME:=% PART_RENAME_VAR FILE_NAME F_2
	
	
	StringUpper, FILE_NAME, FILE_NAME
	
	ControlSetText, Edit1,, ahk_id %HWND_1%
	Control, EditPaste, %FILE_NAME%, Edit1, ahk_id %HWND_1%

	WinGet, HWND_2, ID, A
	IF HWND_2<>%HWND_1%
	{
		#WinActivateForce, ahk_id %HWND_1%
	}
	
	LOOP
	{
		SLEEP 100
		ControlGetPos, x, y, w, h, OK, ahk_id %HWND_1%
		MouseMove, X+10, Y+10

		ControlClick, OK, ahk_id %HWND_1%
		IfWinNotExist, ahk_id %HWND_1%
			BREAK
	}
	

	MouseMove, XPOS, YPOS
	GOSUB DISPLAY_TOOLTIP
	SOUNDBEEP 3000,80
	XPOS=
}
RETURN



RENAME_PATH_OF_JOBS_LEFT_OR_RIGHT:
{

	WinGet, HWND_1, ID, Left Folder ahk_class #32770
	IF !HWND_1 
		WinGet, HWND_1, ID, Right Folder ahk_class #32770

	
	IF !HWND_1
	{
		O_ID=0
		RETURN
	}
	
	; TOOLTIP % HWND_1
	
	; IF HWND_1=%O_ID%
		; RETURN
	; O_ID:=HWND_1


	SLEEP 400

	ControlGettext, OutputVar_2, Edit1, ahk_id %HWND_1%
	IF !OutputVar_2
		RETURN
	IF INSTR(OutputVar_2,"=4_CLOUD_2TB_01:")>0
		RETURN
	IF INSTR(OutputVar_2,"=4_CLOUD_2TB")=0
		RETURN

		
	SET_HO=FALSE
	IF INSTR(OutputVar_2,"=4_CLOUD_2TB:")
	SET_HO=TRUE
	IF INSTR(OutputVar_2,"=4_CLOUD_2TB_1:")
	SET_HO=TRUE
	
	SET_NEXT_EVENT=FALSE
	IF SET_HO=TRUE
	{
		StringReplace, OutputVar_2, OutputVar_2,CLOUD_2TB,CLOUD_2TB_01, All
		StringReplace, OutputVar_2, OutputVar_2,CLOUD_2TB_1,CLOUD_2TB_01, All
		ControlSetText, Edit1,, ahk_id %HWND_1%
		SOUNDBEEP 3000,80
		
		LOOP
		{
			SLEEP 400
			Control, EditPaste, %OutputVar_2%, Edit1, ahk_id %HWND_1%
			;TOOLTIP %OutputVar_3%		
			SOUNDBEEP 3000,50

			ControlGettext, OutputVar_3, Edit1, ahk_id %HWND_1%
			
			; TOOLTIP % "01" OutputVar_2 "`n" "02" OutputVar_3
			IF !OutputVar_2
				BREAK
		
			IF OutputVar_3=%OutputVar_2%
				{
				; MSGBOX % "01" OutputVar_3 "`n" "02" OutputVar_2
				
				BREAK
				}
		SET_NEXT_EVENT=TRUE
		SLEEP 4000
		SOUNDBEEP 3000,40
		}
	}

	
	IF SET_NEXT_EVENT=FALSE
		RETURN 
		
	MSGBOX "HH"

		
	WinGet, HWND_2, ID, A
	IF HWND_2<>%HWND_1%
	{
		#WinActivateForce, ahk_id %HWND_1%
	}
	
	LOOP
	{
		MSGBOX "HH"
		ControlGetPos, X_3, Y_3, , , Button3, ahk_id %HWND_1%
		MouseMove, X_3+10, Y_3+10
		
		ControlClick, Button3, ahk_id %HWND_1%
		IfWinNotExist, ahk_id %HWND_1%
			BREAK
		SOUNDBEEP 3000,80
	}
	

}
RETURN



AUTO_CLONE_JOB:
{

	WinGetTitle, Title, A
	
	POS_VAR:=INSTR(Title,"HDD CLOUD GD1TB")
	IF POS_VAR>0 
	{
		SOUNDBEEP 5000,100
		SLEEP 500
		SENDINPUT !C
	}

	WinGet, HWND_1, ID, Clone Job ahk_class #32770

	IF !HWND_1
	{
		O_ID=0
		RETURN
	}
	
	IF HWND_1=%O_ID%
		RETURN
	
	O_ID:=HWND_1

	SOUNDBEEP 3000,100
		
	LOOP
	{
		SLEEP 100
		ControlGetPos, x, y, w, h, OK, ahk_id %HWND_1%
		MouseMove, X+10, Y+10
		ControlClick, OK, ahk_id %HWND_1%
		IfWinNotExist,  ahk_id %HWND_1%
			BREAK
	}

	SOUNDBEEP 4500,100
}
RETURN


; STRIP SOMETHING FROM RENAME JOB
;; F6::
{
	WinGet, HWND_1, ID, Rename Job ahk_class #32770

	IF !HWND_1
	{
		O_ID=0
		RETURN
	}
	
	
	IF HWND_1=%O_ID%
	{
		RETURN
	}
	
	O_ID:=HWND_1

	
	ControlGettext, OutputVar_1, Button5, ahk_id %HWND_1%
	
	ControlGettext, OutputVar_2, Edit1, ahk_id %HWND_1%
	
	POS_VAR:=INSTR(OutputVar_2," _ 7G")

	OutputVar_4:=SubStr(OutputVar_2,1, POS_VAR)

	SOUNDBEEP 3000,100
	
	IF OutputVar_4
	{
	
		ControlSetText, Edit1,, ahk_id %HWND_1%
		Control, EditPaste, %OutputVar_4%, Edit1, ahk_id %HWND_1%
		OutputVar_4=
		SOUNDBEEP 4000,100
		
		SLEEP 100
		LOOP
		{
			SLEEP 100
			ControlClick, OK, ahk_id %HWND_1%
			IfWinNotExist,  ahk_id %HWND_1%
				BREAK
		}

		SOUNDBEEP 4500,100
	}
}
RETURN


; \\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_04_4_SAMSUNG_4TB_HUBIC\D4\#
; =4_SAMSUNG_4TB_HUBIC:\D4\#
;;F6::
{
	WinGet, HWND_1, ID, Right Folder ahk_class #32770

	ControlGettext, OutputVar_1, Button5, ahk_id %HWND_1%
	
	ControlGettext, OutputVar_2, Edit1, ahk_id %HWND_1%

	POS_VAR:=INSTR(OutputVar_2,"4TB_HUBIC\")
	
	OutputVar_4:=SubStr(OutputVar_2, POS_VAR+10)

	
	OutputVar_4="=4_SAMSUNG_4TB_HUBIC:\"%OutputVar_4%
	StringReplace, OutputVar_4, OutputVar_4,/,\, All
	NV := chr(34)
	StringReplace, OutputVar_4, OutputVar_4,%NV%,, All
	
	; TOOLTIP % OutputVar_4
	; MSGBOX % OutputVar_4
	; RETURN
	SOUNDBEEP 3000,100
	
	
	IF OutputVar_4
	{
	
		ControlSetText, Edit1,,  ahk_id %HWND_1%
		Control, EditPaste, %OutputVar_4%, Edit1, ahk_id %HWND_1%
		OutputVar_4=
		SOUNDBEEP 4000,100
		
		#WinActivateForce, ahk_id %HWND_1%
		SLEEP 100
		WinWAIT, ahk_id %HWND_1%
		SLEEP 100
		SENDINPUT {ENTER}
		SLEEP 100
		LOOP
		{
			SLEEP 100
			ControlClick, OK, ahk_id %HWND_1%
			IfWinNotExist,  ahk_id %HWND_1%
				BREAK
		}

		SOUNDBEEP 4500,100
	}
}
RETURN


;; F6::
{

	WinGet, HWND_1, ID, Right Folder ahk_class #32770

	ControlGettext, OutputVar_1, Button5, ahk_id %HWND_1%
	
	IF INSTR(OutputVar_1,"Secure Mode (Encrypted by TLS)")
	{
		ControlGettext, OutputVar_2, Edit1, ahk_id %HWND_1%

		POS_VAR:=INSTR(OutputVar_2,"goodsync/")
		
		OutputVar_4:=SubStr(OutputVar_2, POS_VAR+10)

		; O:\C\0 00 LINK SET QUICKER MOVER
		
		OutputVar_4="=4_SAMSUNG_4TB_HUBIC"%OutputVar_4%
		StringReplace, OutputVar_4, OutputVar_4,/,\, All
		NV := chr(34)
		StringReplace, OutputVar_4, OutputVar_4,%NV%,, All
		
		;TOOLTIP % OutputVar_4
		SOUNDBEEP 3000,100
	}
	
	
	IF OutputVar_4
	{
	
		ControlSetText, Edit1,,  ahk_id %HWND_1%
		Control, EditPaste, %OutputVar_4%, Edit1, ahk_id %HWND_1%
		OutputVar_4=
		SOUNDBEEP 4000,100
		
		#WinActivateForce, ahk_id %HWND_1%
		SLEEP 100
		WinWAIT, ahk_id %HWND_1%
		SLEEP 100
		SENDINPUT {ENTER}
		SLEEP 100
		LOOP
		{
			SLEEP 100
			ControlClick, OK, ahk_id %HWND_1%
			IfWinNotExist,  ahk_id %HWND_1%
				BREAK
		}

		SOUNDBEEP 4500,100
	}
}
RETURN





; F5::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 500

	; VAR_CLIP=%clipboard%

	; ; StringUpper, VAR_CLIP, VAR_CLIP
	
	; StringReplace, VAR_CLIP, VAR_CLIP,/file://,, All
	
	; Sleep 800
	; clipboard=%VAR_CLIP%

	; Sleep 400
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}V
	; Send {ctrl up}

	; ; SLEEP 500
	; ; LOOP
	; ; {
		; ; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; ; SLEEP 100
		; ; ControlClick, OK, Rename ahk_class #32770
		; ; IfWinNotExist, Rename ahk_class #32770
			; ; BREAK
	; ; }

; }
; RETURN


; F5::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {right}
	; Sleep 100
	; Send {ctrl down}V
	; Send {ctrl up}

	; SLEEP 500
	; LOOP
	; {
		; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; SLEEP 100
		; ControlClick, OK, Folder ahk_class #32770
		; IfWinNotExist, Folder ahk_class #32770
			; BREAK
	; }
; }
; RETURN


; F4::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 500

	; VAR_CLIP=%clipboard%

	; ;StringUpper, VAR_CLIP, VAR_CLIP

	; ;gstps://nas-qnap-ml.matt-lan-2.goodsync/file:///share/CACHEDEV1_DATA/Multimedia/SW/SW SOFTWARE
	; ;\\NAS-QNAP-ML\PUBLIC_CONTACT
	
	; ;StringReplace, VAR_CLIP, VAR_CLIP,`:`\,%A_Space%, All
	; ;StringReplace, VAR_CLIP, VAR_CLIP,`\`\NAS-QNAP-ML,gstps`://nas-qnap-ml`.matt-lan-2`.goodsync/file`:///share/CACHEDEV1_DATA, All
	; StringReplace, VAR_CLIP, VAR_CLIP,`\`\NAS-QNAP-ML,, All

	; Sleep 800
	; clipboard=%VAR_CLIP%

	; ; Sleep 400
	; ; Send {ctrl down}A
	; ; Sleep 100
	; ; Send {ctrl up}
	; ; Sleep 100
	; ; Send {ctrl down}V
	; ; Send {ctrl up}

	; ; SLEEP 500
	; ; LOOP
	; ; {
		; ; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; ; SLEEP 100
		; ; ControlClick, OK, Folder ahk_class #32770
		; ; IfWinNotExist, Folder ahk_class #32770
			; ; BREAK
	; ; }

; }
; RETURN

; F4::
; {
	; DetectHiddenWindows, off
	; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {delete}

	; SLEEP 200
	; LOOP
	; {
		; SLEEP 200
		; #WinActivateForce, Folder ahk_class #32770
        ; ;SoundBeep , 2000 , 400
		; ControlClick, OK, Folder ahk_class #32770
		; IfWinNotExist, Folder ahk_class #32770
			; BREAK
	; }
; }
; RETURN




; F4::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 500

	; VAR_CLIP=%clipboard%
	; StringUpper, VAR_CLIP, VAR_CLIP

	; StringReplace, VAR_CLIP, VAR_CLIP,`:`\,%A_Space%, All
	; StringReplace, VAR_CLIP, VAR_CLIP,`\,_, All
	; StringReplace, VAR_CLIP, VAR_CLIP,`.,_, All
	; Sleep 800
	; clipboard=%VAR_CLIP%

	; ; Sleep 400
	; ; Send {ctrl down}A
	; ; Sleep 100
	; ; Send {ctrl up}
	; ; Sleep 100
	; ; Send {ctrl down}V
	; ; Send {ctrl up}

	; SLEEP 500
	; LOOP
	; {
		; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; SLEEP 100
		; ControlClick, OK, Folder ahk_class #32770
		; IfWinNotExist, Folder ahk_class #32770
			; BREAK
	; }

	; }
; RETURN



; F4::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 500

	; VAR_CLIP=%clipboard%
	; StringUpper, VAR_CLIP, VAR_CLIP

	; StringReplace, VAR_CLIP, VAR_CLIP,vs LINKFOLDER,, All
	; Sleep 500
	; clipboard=%VAR_CLIP%

	; Sleep 400
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}V
	; Send {ctrl up}

	; SLEEP 500
	; LOOP
	; {
		; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; SLEEP 100
		; ControlClick, OK, Rename Job ahk_class #32770
		; IfWinNotExist, Rename Job ahk_class #32770
			; BREAK
	; }

	; }
; RETURN




; F4::
; {
	; Send {F2}
	; Sleep 100
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 500

	; VAR_CLIP=%clipboard%
	; StringUpper, VAR_CLIP, VAR_CLIP
	; StringReplace, VAR_CLIP, VAR_CLIP,S01,S02, All
	; StringReplace, VAR_CLIP, VAR_CLIP, (1),, All
	; Sleep 500
	; clipboard=%VAR_CLIP%

	; Sleep 400
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}V
	; Send {ctrl up}

	; SLEEP 500
	; LOOP
	; {
		; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; SLEEP 100
		; ControlClick, OK, Rename Job ahk_class #32770
		; IfWinNotExist, Rename Job ahk_class #32770
			; BREAK
	; }

	; }
; RETURN


; F4::
; {
	; Send {ctrl down}A
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {ctrl down}C
	; Send {ctrl up}
	; Sleep 200
	; VAR_CLIP=%clipboard%
	; StringUpper, VAR_CLIP, VAR_CLIP
	; StringReplace, VAR_CLIP, VAR_CLIP, D`:`\, , All
	; Sleep 200
	; clipboard=%VAR_CLIP%
; }
; RETURN



; F4::
; {
	; IF GetKeyState("Capslock", "T") 
		; Send {shift down}
	; SENDINPUT C{:}
	; IF GetKeyState("Capslock", "T") 
		; Send {shift up}
	; SENDINPUT {ENTER}
	; SLEEP 500
	; LOOP
	; {
		; SetTitleMatchMode 2  ; Avoids specify the full path of the file below.
		; SLEEP 100
		; ControlClick, OK, Folder ahk_class #32770
		; IfWinNotExist, Folder ahk_class #32770
			; BREAK
	; }
; }
; RETURN



; F4::
; {
	; Send {ctrl down}a
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; SEND 12
	; SENDINPUT {ENTER}
; }
; RETURN



; F4::
; {
	; Send {ctrl down}a
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {shift down}
	; Send C:
	; Send {shift up}
	; Send \
	; Send {shift down}
	; Send PS
	; Send {shift up}
	; Send tart
	; Sendinput {ENTER}
	; SLEEP 500
	; LOOP
	; {
	; SLEEP 100
	; ControlClick, OK, Left Folder ahk_class #32770
	; IfWinNotExist, Left Folder ahk_class #32770
		; BREAK
	; }
; }
; Return

; F4::
; {
	; SEND 0
	; SEND {SPACE}
	; SENDINPUT {ENTER}
; }
; RETURN


; F4::
; {
	; Send 7
	; Send {shift down}
	; Sleep 100
	; Sendinput G
	; Send {shift up}
	; Sendinput {ENTER}
; }
; Return


; F4::
; {
	; Send {ctrl down}a
	; Sleep 100
	; Send {ctrl up}
	; Sleep 100
	; Send {shift down}
	; Send C:
	; Send {shift up}
	; Send \
	; Send {shift down}
	; Send SCRIPTER
	; Send {shift up}
; }
; Return



;F4::
;{
;Sendinput existboth
;}
;Return

;-------------------------------------------------------------------------
; CLIPBOARD CTRL 'A'-ALL AND THEN COPY
;-------------------------------------------------------------------------
;F4::
;   Send {ctrl down}a
;   Sleep 100
;   Send {ctrl up}
;   Sleep 100
;   Send {ctrl down}c{ctrl up}
;Return
;-------------------------------------------------------------------------

;F4::
;	Send Mill View Hospital
;Return


;F1::
;   Send {ctrl down}A{ctrl up}
;   Sleep 100
;   Send {ctrl down}V{ctrl up}
;Return
;-------------------------------------------------------------------------

;lock control key up or down shift and ctrl
;WHEN SELECT IS DIFFERCULT AS CTRL BOUCE UP WHEN HOLD KEY DOWN
;-------------------------------------------------------------------------
;+F4::
;   Send {ctrl down}
;   SoundBeep , 1500 , 400
;Return
;-------------------------------------------------------------------------
;^F4::
;   Send {ctrl up}
;   SoundBeep , 2000 , 400
;Return

;NORTON
;F4::
;   SEND :\
;Return 

;-------------------------------------------------------------------------
;NOT AT MOMENT
;GRAB A LINE BEGIN TO END
;F4::
;   Send {home}+{end}{ctrl down}c{ctrl up}
;-------------------------------------------------------------------------

;-------------------------------------------------------------------------
; F4 __ CLIPBOARD CTRL 'A'-ALL DELETE AND WAIT AND THEN PASTE CTRL V 
;-------------------------------------------------------------------------
;F4::
;   SoundBeep , 1500 , 400
;   Sendinput ^a
;   Sendinput {delete}
;   Sendinput {ctrl down}
;   sleep 200
;   Sendinput v
;   Sendinput {ctrl up}
;-------------------------------------------------------------------------


;F4::
;   Sendinput ^a
;   Sendinput {delete}
;   Sendinput 48
;   Send,{ENTER}
;   SoundBeep , 1500 , 400


;------------------------------------------------------------
;----
;Deselect All Text - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/122592-deselect-all-text/
;----
;GEV
;Members
;1364 posts
;Last active: Mar 15 2017 06:36 AM
;Joined: 23 Oct 2013
;Alt+Click doesn't open a link. Try this:
;------------------------------------------------------------
;F4::
;WinGetPos,,,Xmax,Ymax,A ; get active window size
;Ycenter := Ymax/2
;Send, {ALTDOWN}
;ControlClick, x10 y%Ycenter%, A   ;this is the safest point, I think
;Send, {ALTUP}
;return


;--------------------------------------------------------------------



;-------------------------------------------------------------
;MAKE WIN + Y WORK TO DO A SCROLL LOCK FOR WINDOWSE STOP CLEAR
;-------------------------------------------------------------
;NOT WORKING TO DO A WIN + KEY TO RELACE
;LIKE AUTOKEY USE WIN + A -- TO STOP A SPY
;-------------------------------------------------------------
;^::

;#Y::^L

;#Y::^ScrollLock


;#Y::
;SEND ^L
;SEND ScrollLock
;-------------------------------------------------------------


;---------------------------------------------------
;CTRL F FIND
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}
;   Send,{ctrl down}f{ctrl up}
;---------------------------------------------------
;CTRL F FIND AND ENTER ON HIGH-LIGHTED SELECTED TEXT
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}




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

