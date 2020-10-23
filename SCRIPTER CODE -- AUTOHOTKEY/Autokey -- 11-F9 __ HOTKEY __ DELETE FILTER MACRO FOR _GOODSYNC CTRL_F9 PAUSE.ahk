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

;
;--------------------
#SingleInstance force
;--------------------
;--------------------

;---------------------------------------------------------
; CODE INITIALIZE SOUND EFFECT LEARN
;---------------------------------------------------------
SoundBeep , 1500 , 400
;---------------------------------------------------------
Suspend On
PAUSE
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
;-------------------------------------------------------------------------
;F10::^F10
;F10::
;SoundBeep , 1500 , 400
;Sendinput +{PrintScreen}
;Sendinput ^{F10}
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------
;-------------------------------------------------------------------------


; F9::Sendinput {Enter}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Del}{Enter}{Up}

^F9::
SoundBeep , 1500 , 400
Suspend On
PAUSE
return

Return

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
