#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;---------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CRTL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
;---------------------------------------------------------
;WANT COPY TEXT AND REPEAT PASTE IT DOWN A LINE HOM DOWN PASTE PUT REMARK IN
;---------------------------------------------------------

;
;--------------------
#SingleInstance force
;--------------------
;--------------------

;CLIPBOARD A ALL AND THEN COPY
;F4::
;   Send {ctrl down}a
;   Sleep 100
;   Send {ctrl up}
;   Sleep 100
;   Send {ctrl down}c{ctrl up}
;Return


;GRAB A LINE BEGIN TO END
;F4::
;   Send {home}+{end}{ctrl down}c{ctrl up}

F4::
{    
    Sendinput ^a
    ;sleep 200
    Sendinput {delete}
    ;sleep 200
    ;Sendinput ^v
    Sendinput {ctrl down}
    sleep 200
    Sendinput v{ctrl up}
}

;-------------------------------------------------------------------------
;REPLACE F10 TO DO CONTROL PRINT SCREEN
;FOR CLIPBOARD SCREEN SHOT -- CODE PICPICK WON'T HOT KEY F10 ON WIN 7 AND UP
;-------------------------------------------------------------------------
F10::
Sendinput +{PrintScreen}
;Shift +

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
