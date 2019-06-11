#NoEnv 
; #Warn
SendMode Input 
SetWorkingDir %A_ScriptDir% 

#SingleInstance force
#Persistent

SetStoreCapslockMode, off

SoundBeep , 1500 , 400
RETURN

; -------------------------------------------------------------------
; THIS METHOD DO THE CASE INSENSITIVE PROPER
; ADD THEN ASTERISK AND NONE REQUIRE SPACE BEFORE OR AFTER 
; ASTERISK STILL REQUIRE SPACE BEFORE
; TRY ?
; USE BOTH LIKE HERE
; THIS WILL NOT CASE PROPER IF NOT USE -- SetStoreCapslockMode, off
; -------------------------------------------------------------------
:*?:n`;t::
	SENDINPUT n't
RETURN







; :*:hima::
; MESSENGER_KEY=Hi Marianne and Eddie
; GOSUB STRING_INVERT_MESSENGER
; SENDINPUT %MESSENGER_KEY%
; RETURN


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
; THIS WILL NOT CASE PROPER IF NOT USE -- SetStoreCapslockMode, off
; -------------------------------------------------------------------
; :*?:n`;t::
	; SENDINPUT n't
; RETURN


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
		TOOLTIP % YES ON
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

