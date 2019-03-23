#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 07-WINAMP VOLUME HOTKEY CHANGER.AHK

; I HAVE A CONVERTOR WHEN I TYPE I DO ALL IN CAP LOCK 
; I MAYBE DO IT LATER
; TOO MUCH TROUBLE TIME THIS LOT TAKEN
; JUST A MATTER OF CLIPBOARD IT OUT IN CHUNK

; READ IT THROUGH ALL IS EXPLAIN
; IT IS MY DELIGHT AND IT'S A PLEASURE TO PRESENT THIS CODE 
; AS FREE FOR EDUCATION 
; WORK MY HEAD DOING THIS -- ENJOY THE CODE -- TROUBLE IT TAKEN INTO 

; AUTO KEYS IS DIFFICULT FOR BEGINNER OF MY STYLE
; EVERYTHING IS OF ASSUMPTION OF WILD CARD TYPE ITEM
; AND ANOTHER
; THE FORUM SEEM LIKE ANOTHER ANGLE ON THE HELP FILE
; BUT AND ONE IS INTERNET -- HELP FILE IS NOT AND QUICKER SLOWER MACHINE WITH LESS QUICK


; -----------------------
; NEW CODE WINAMP CHANGER
; -----------------------
; 03:29 PM -- TUE 01 NOV

; 19:31 -- PROJECT ANOTHER GO
; TO
; 18:27 -- RESUME WORK IN TWO PROJECT SPLIT TWO PART
; 20:10 -- DONE
; 23:17 -- GOING ON-LINE GRAMMAR
; 23:23 -- GOING ON-LINE TYPING -- DONE
 
;---------------------------------------------------------------- 
; TIME TO SOFTWARE IF THE PRICE OF HARDWARE IS HIGH AT END OF YEAR
;---------------------------------------------------------------- 

;---------------------------------------------------------
; WINAMP HOTKEY MUST BE ENABLED BOTTOM CHECK BOX 
; NOT BOTH TOP ONE MEDIA KEY
; FOR NUMBER PAD FUNCTION HOTKEY
;
; NORMAL ARROW KEY UP DOWN DO PROBLEM WITH SENDKEY DOING CAP LOCK OFF ON
;---------------------------------------------------------
;---------------------------------------------------------
; PROJECT 01 ---- PAGE UP AND PAGE DOWN -- VOL CHANGE
; HERE
;---------------------------------------------------------
; PROJECT 02 ---- VOL CHANGE ON TRACK WINAMP CHANGE --
;---------------------------------------------------------
; PROJECT 02 IS THE PARTNER OF PROJECT 01
; REQUIRED SEPARATE
; PROJECT 02 LEARN SUBROUTINE CALL
; PROJECT 02 LEARN LOOP TO A SUBROUTINE ON TIMER EVENT
;---------------------------------------------------------

;--------------------
#SingleInstance force
;--------------------
; UP WITH CONTROL PAGER UP HARDER WITH CONTROL
; NICEST PLACE ON MY KEYBOARD
; JOYPAD ON THE CONTROL MODULE LATER
;--------------------
*Pgup::
	IfWinExist ahk_class Winamp v1.x
	{
		; -- Not in notepad 
		; -- SCRIPT FILTER ITEM NOT USE-- WHILE WORKER
		IfWinNotActive ahk_class Notepad
		IfWinNotActive ahk_class Notepad++
		{
		{
			SendMessage, 1024, 0, 104   ; IsWinampPlaying?
			WinampPlayStatus := ErrorLevel
			If( WinampPlayStatus == 1 )
			{
				;------------------------------------------------------------
				; SET MASTER VOLUME OF WINAMP -- RATHER THAN STEP UP DOWN OR WITH KEY
				;------------------------------------------------------------
				; 255 IS THE VOLUME LEVEL UP
				; VOLUME ONLY UP IF PLAYING
				; DON'T WANT ON HEART ATTACK
				;------------------------------------------------------------		
				SendMessage, 1024, Mod( 255, 256 ), 122
				SOUNDBEEP
			}
		}
		}
	}
Return

;--------------------
; DOWN WITH PAGE DOWN -- EASY
; PAGE DOWN WITHOUT CONTROL CODE KEY WON'T WORK FOR ME
; LATER HUNGRY THAT ONE FOOD
; STAR SOLUTION
; CHANGE ONE TO START WORKING SOLUTION AND THEN THE OTHER DON'T WORK
; UNLESS BOTH SAME -- ANSWER WORK AROUND WORKING OVER TIME -- LATER
;--------------------
IfWinNotActive Notepad
*Pgdn::
	IfWinExist ahk_class Winamp v1.x
	{
		; -- Not in notepad 
		; -- SCRIPT FILTER ITEM NOT USE-- WHILE WORKER
		IfWinNotActive ahk_class Notepad
		IfWinNotActive ahk_class Notepad++
		{
		{
			;------------------------------------------------------------
			; SET MASTER VOLUME OF WINAMP -- RATHER THAN STEP UP DOWN OR WITH KEY
			;------------------------------------------------------------
			; 0 IS THE VOLUME LEVEL DOWN
			; NOT BOTHER BY WINAMP PAUSED OR PLAYING
			;------------------------------------------------------------		
			SendMessage, 1024, Mod( 1, 256 ), 122
			SOUNDBEEP
		}
		}
	}
Return

;--------------------------------------
; JUNK REMARK LINE AREA

; ADVANCED CLASS SOFTWARE -- RESULT I WANTED HWND IN NUMERIC

; Retrieve the ID/HWND of the active window
; id := WinExist("A")
; MsgBox % id


; -------------------------------------------------------
; MASTER VOLUME SOURCE CODE FIND HERE CREDIT IDIOT TALKER
; -------------------------------------------------------
;----
; Winamp: set Volume to a given value - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/53975-winamp-set-volume-to-a-given-value/
;----
;----
; SETVOLUME WINAMP AUTOHOTKEYS HOOKER - Google Search
; https://www.google.co.uk/search?q=SETVOLUME+WINAMP+AUTOHOTKEYS HOOKER&rlz=1C1SKPL_enGB417&oq=SETVOLUME+WINAMP+AUTOHOTKEYS HOOKER&aqs=chrome..69i57.23407j0j7&sourceid=chrome&ie=UTF-8
; ----
; -------------------------------------------------------

; -------------------------------------------------------
;Send,{ctrl down}{alt down}{down}{alt up}{ctrl up}
;Send,{^!{up}} ;CTRL ALT
;Send,{Numpad2}  ; volume D
;Send,{Numpad8}  ; volume up
