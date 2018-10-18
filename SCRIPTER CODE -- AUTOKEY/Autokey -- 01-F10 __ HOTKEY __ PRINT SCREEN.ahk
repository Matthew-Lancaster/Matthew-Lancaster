;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
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
;

; -------------------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
; -------------------------------------------------------------------
;WANT COPY TEXT AND REPEAT PASTE IT DOWN A LINE HOME DOWN PASTE PUT REMARK IN
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2001-F10%20__%20HOTKEY%20__%20PRINT%20SCREEN.ahk
; ----
;# ------------------------------------------------------------------


;
;--------------------
#SingleInstance force
;--------------------
;--------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1
SETTIMER TIMER_ENTER,OFF
DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path
; -------------------------------------------------------------------
; ENTER THE COUNTER BEGIN NUMBER FOR FACEBOOK PHOTO DESCRIPTION 
; AT THE NUMBER NEXT NEEDER TO BE ENTER
VAR_COUNTER=351
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; GLOBAL FILE_SCRIPT
; GLOBAL FILE_SCRIPT_COUNT

FILE_SCRIPT_COUNT=0
FILE_SCRIPT := Object()
; FILE_SCRIPT := []

FILE_PATH_WILDPATH_JPG=D:\DSC\2015-2018\2018 CyberShot HX60V\WORK\*.JPG

Loop, Files, %FILE_PATH_WILDPATH_JPG%
{
	FILE_SCRIPT[A_Index] := A_LoopFileName
	FILE_SCRIPT_COUNT := A_Index
}	

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

; USE 0 FOR START SEARCH LOCATION DOES LIKE VB INSTRREV IN REVERSE SEARCH  
; -----------------------------------------------------------------------
SET_String:=SubStr(FILE_PATH_WILDPATH_JPG,1,InStr(FILE_PATH_WILDPATH_JPG,"\",,0)-1)

; WHEN CREATE ALBUM FIRST TIME PAGE
FACEBOOK_URL_TITLE_1=Matthew Lancaster - Google Chrome
; WHEN IN THE ALBUMS EDIT PAGE
FACEBOOK_URL_TITLE_2=Facebook - Google Chrome

IfWinExist, %SET_String%
	SET_GO=FALSE
	IfWinExist, %FACEBOOK_URL_TITLE_1%
	{
		SET_GO=TRUE
		SETTIMER F4,3000
	}
	IfWinExist, %FACEBOOK_URL_TITLE_2%
	{
		SET_GO=TRUE
		; WHEN IN THE ALBUMS EDIT PAGE _ TIMER FOR UPDATE HAS 
		; TO BE LESS QUICK OR ERROR TAB TO NEXT ONE
		; GIVING ENOUGH SPEED IS CRITICAL TAB WILL JUMP IF NOT READY FOR NEXT ONE IN
		FACEBOOK_TIMER_DELAY=14000
		SETTIMER F4,%FACEBOOK_TIMER_DELAY%
	}
	IF SET_GO=TRUE
	{
		WinActivate ; use the window found above

	}
Return


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

F12::Process,Close,WScript.exe
; C:\Windows\SysWOW64\WScript.exe


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



F4::
{

FILE_NAME := % FILE_SCRIPT[VAR_COUNTER]

IF GetKeyState("Capslock", "T")
{
	StringLower, FILE_NAME, FILE_NAME
}
ELSE
{
	StringUpper, FILE_NAME, FILE_NAME
}
	
POS_VAR:=InStr(FILE_NAME,"_")

if POS_VAR>0 
{
	POS_VAR-=1
	OutputVar_1:=SubStr(FILE_NAME, 1, POS_VAR)
	POS_VAR+=2
	OutputVar_2:=SubStr(FILE_NAME, POS_VAR)
}
else
{
	OutputVar_1=%FILE_NAME%
}

StringReplace, OutputVar_2, OutputVar_2,.JPG,,

SoundBeep , 1000 , 150

IfWinExist, %SET_String%
{
	IfWinExist, %FACEBOOK_URL_TITLE_1%
	{
		WinActivate ; use the window found above
	}
	IfWinExist, %FACEBOOK_URL_TITLE_2%
	{
		WinActivate ; use the window found above
	}
}


Sendinput ^a{delete}
Sendinput %VAR_COUNTER% of %FILE_SCRIPT_COUNT%`n
Sendinput %OutputVar_1%`n

if POS_VAR>0 
	Sendinput %OutputVar_2%`n

; -------------------------------------------------------------------
; IF ALREADY PUBLISHED PHOTO AND WANT TO ADD INFO DESCRIPTION 
; USE ONE TAB INSTEAD 
; -------------------------------------------------------------------

IfWinExist, %FACEBOOK_URL_TITLE_2%
{
	Sendinput ^{home}
	Sendinput {tab}
	; IMPORTANT TO HAVE BIGGER SLEEP AFTER TAB NEXT ITEM OR IT LAND 
	; ON WRONG LOCATION
	; ---------------------------------------------------------------
	SLEEP 1500
	Sendinput ^{home}
}
	
IfWinExist, %FACEBOOK_URL_TITLE_1%
{
	Sendinput ^{home}
	Sendinput {tab}{tab}{tab}{tab}{tab}
	SLEEP 1000
}


SETTIMER F4,off
SLEEP 200
SETTIMER F4,%FACEBOOK_TIMER_DELAY%


VAR_COUNTER+=1
if VAR_COUNTER>%FILE_SCRIPT_COUNT%
{
	SETTIMER F4,off
	SoundBeep , 1000 , 150
	SoundBeep , 2000 , 200
	SoundBeep , 1000 , 150

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


TIMER_ENTER:
	DetectHiddenWindows, oFF
	SetTitleMatchMode 3  ; Specify Full path.
	IfWinExist, Clone Job ahk_class #32770
	{
		ControlClick, OK, Clone Job ahk_class #32770
		SoundBeep , 2500 , 100
	}
RETURN

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

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

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

