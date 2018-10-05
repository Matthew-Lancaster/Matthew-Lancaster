;  =============================================================
;  =============================================================


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;---------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
;---------------------------------------------------------
;WANT COPY TEXT AND REPEAT PASTE IT DOWN A LINE HOME DOWN PASTE PUT REMARK IN
;---------------------------------------------------------

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
; https://www.dropbox.com/s/3t0ei0xl5k8r0la/Autokey%20--%2001-F10%20__%20HOTKEY%20__%20PRINT%20SCREEN.ahk?dl=0
;# ------------------------------------------------------------------


;
;--------------------
#SingleInstance force
;--------------------
;--------------------

;---------------------------------------------------------
; CODE INITIALIZE
;---------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1
SETTIMER TIMER_ENTER,OFF
DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path
; -------------------------------------------------------------------
VAR_COUNTER=1
; -------------------------------------------------------------------

; GLOBAL FILE_SCRIPT
; GLOBAL FILE_SCRIPT_COUNT

FILE_SCRIPT_COUNT=0
FILE_SCRIPT := Object()
; FILE_SCRIPT := []

Loop, Files, D:\DSC\2015-2018\2018 CyberShot HX60V\DCIM\2018 09 13 _ WALK BOTH TESCO HOVE ALSO _ ALL WAY ALONG SEAFRONT & BACK TURNING UP A BIT TO TOWN\*.JPG
{
	FILE_SCRIPT[A_Index] := A_LoopFileName
	FILE_SCRIPT_COUNT := A_Index
}	
	
Return



;-------------------------------------------------------------------------
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
;-------------------------------------------------------------------------
;F10::^F10

F10::
	SoundBeep , 1500 , 400
	;SetCapsLockState , Off
	;SetNumLockState , Off
	;SetScrollLockState , Off

	Sendinput ^{F10}
	; LOGITECH MOUSE INFO POP UP WONT LIKE CAPS CHANGE OR HOT KEY F10 THINK INFO REQUIRE ABOUT CAPITAL CHANGE OR TURN LOGITECH NOTIFY OFF
Return
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------


 
; F4::
; {

; FILE_NAME := % FILE_SCRIPT[VAR_COUNTER]

; IF GetKeyState("Capslock", "T")
; {
	; StringLower, FILE_NAME, FILE_NAME
; }
; ELSE
; {
	; StringUpper, FILE_NAME, FILE_NAME
; }

	
; POS_VAR:=InStr(FILE_NAME,"_")

; ; MSGBOX % FILE_NAME

; if POS_VAR>0 
; {
	; POS_VAR-=1
	; OutputVar_1:=SubStr(FILE_NAME, 1, POS_VAR)
	; POS_VAR+=2
	; OutputVar_2:=SubStr(FILE_NAME, POS_VAR)
; }
; else
; {
	; OutputVar_1=%FILE_NAME%
; }

; StringReplace, OutputVar_2, OutputVar_2,.JPG,,

; Send %VAR_COUNTER% of %FILE_SCRIPT_COUNT%`n
; Send %OutputVar_1%`n
; if POS_VAR>0 
	; Send %OutputVar_2%`n

; ; Send {tab}
; Send {tab}{tab}{tab}{tab}{tab}

; VAR_COUNTER+=1
; }



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

