#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CRTL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
;
;--------------------
#SingleInstance force
;--------------------
;--------------------
F4::
   Send,{ctrl down}a{ctrl up}{ctrl down}c{ctrl up}
;-----------
;CTRL F FIND
;-----------
;F6::
;   Send,{ctrl down}f{ctrl up}
;---------------------------------------------------
;CTRL F FIND AND ENTER ON HIGH-LIGHTED SELECTED TEXT
;---------------------------------------------------
F6::
   Send,{ctrl down}f{ctrl up}{ENTER}
Return