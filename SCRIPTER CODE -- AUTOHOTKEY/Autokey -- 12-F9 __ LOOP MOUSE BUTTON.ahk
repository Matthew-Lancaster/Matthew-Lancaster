#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;---------------------------------------------------------
;WANT 
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
SUSPEND
PAUSE


;----
;AUTOHOTKEYS REPEAT PRESS MOUSE BUTTON - Google Search
;https://www.google.co.uk/search?q=AUTOHOTKEYS+REPEAT+PRESS+MOUSE+BUTTON&rlz=1C1CHBF_en-GBGB759GB759&oq=AUTOHOTKEYS+REPEAT+PRESS+MOUSE+BUTTON&aqs=chrome..69i57.10732j0j7&sourceid=chrome&ie=UTF-8
;----
;----
;Repeatedly click the mouse button - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/31722-repeatedly-click-the-mouse-button/
;----

;---------------------------------------------------------



^!a::SetTimer, ClickIt, 1
^!d::SetTimer, ClickIt, Off
ClickIt:
Click
Return




;Started = 0

; F9::
;     If (Started == 0)
;     {
;         ; MouseGetPos, X, Y
;          Started = 1
;		  SoundBeep , 1500 , 400
;          SetTimer, Clicking, 1
;     }
;     Else
;     {
;          Started = 0
;		  SoundBeep , 2000 , 400
;          SetTimer, Clicking, Off
;     }
;Return

;Clicking:
;     Click
;;, %X%, %Y%
;Return


^F9::
SoundBeep , 1500 , 400
SUSPEND
return

Return

