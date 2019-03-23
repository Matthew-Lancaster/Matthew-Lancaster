#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;


;NEW CODE WINAMP CHANGER
;03:29 PM TUE 01 NOV


;WINAMP HOT KEY MUST BE ENABLED BOTTOM CHECK BOX 
;NOT BOTH TOP ONE MEDIA KEY
;FOR NUMBER PAD FUNCTION HOT KEY
;
;NORMAL ARROW KEY UP DOWN DO PROBLEM WITH SENDKEY DOING CAP LOCK OFF ON

;TO
;18:27 RESUME WORK IN TWO PROJECT SPLIT TWO PART


;---------------------------------------------------------
;PROJECT 01 ---- HERE ------------------------------------

;PROJECT 02 ---- HERE ------------------------------------
;---------------------------------------------------------
;PROJECT 02 LEARN SUB ROUTINE CALL
;PROJECT 02 LEARN LOOP TO A SUB ROUTINE ON TIMER EVENT
;---------------------------------------------------------

;---------------------------------------------------------
;WANT 
;---------------------------------------------------------

;--------------------
#SingleInstance force
;--------------------
;--------------------
^Pgup::
	X = 0
	SoundBeep

	;--------------------------------------------------------------------
	Loop 
	{
		X += 1
		Send,{Numpad2}  ; volume down

		;sleep 40

		if (X > 100)
		{
			SoundBeep
			break
		}
	}
Return

;--------------------
;--------------------
^Pgdn::
	X = 0
	SoundBeep

	;--------------------------------------------------------------------
	Loop 
	{
		X += 1
		Send,{Numpad8}  ; volume up
		
		;sleep 40

		if (X > 100)
		{
			SoundBeep
			break
		}
	}
Return



;Send,{ctrl down}{alt down}{down}{alt up}{ctrl up}
;Send,{^!{up}} ;CTRL ALT
