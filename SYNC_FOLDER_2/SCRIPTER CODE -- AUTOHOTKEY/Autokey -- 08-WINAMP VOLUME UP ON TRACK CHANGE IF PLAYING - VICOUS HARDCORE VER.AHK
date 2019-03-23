#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;NEW CODE WINAMP CHANGER VOLUME UP ON TRACK CHANGE IF PLAYING ACTIVE
;17:27 PM TUE 01 NOV

;---------------------------------------------------------
;WANT 
;---------------------------------------------------------

;--------------------
#SingleInstance force
;--------------------
;--------------------

;#Persistent
;Run Winamp

;Gui, -Caption
;Winset, Transcolor, 000000 192, wintitle
;Gui, +AlwaysOnTop
;Gui, Color, F50000
;Gui, Show, x1280 y-1 h40 w1280  ; bottom of screen mask

SavedWATitle=
setTimer WinampTitle,100
return

;-----the buttons-----
WinampTitle:
IfWinExist ahk_class Winamp v1.x
{
	WinGetTitle CurrentWATitle,ahk_class Winamp v1.x
	if CurrentWATitle<>%SavedWATitle%
	{
		;Gui 1:-Caption
		;Gui 1:+AlwaysOnTop
		;Gui 1:Color, 000000
		;Gui 1:Submit, nohide
		;Gui 1:Font, S16 cRed Bold , Tahoma
		;Gui 1:Add, Button, x5 y1 w25 h35, <
		;Gui 1:Font, S16 cRed Bold , Tahoma
		;Gui 1:Add, Button, x118 y1 w25 h35, >
		;Gui 1:Font, S16 cRed Bold , Tahoma
		;Gui 1:Add, Button, x29 y1 w90 h35, PLAY
		;Gui 1:Font, S14 cRed Bold , Tahoma
		;Gui 1:Add, Button, x147 y1 w60 h35, STOP
		;Gui 1:Font, S14 cRed Bold , Tahoma
		;Gui 1:Add, Button, x212 y1 w100 h35, VISUALS
		;Gui 1:Font, S9 cRed, Tahoma
		;Gui 1:Add, Button, x320 y1 w45 h16, UP
		;Gui 1:Font, S9 cRed, Tahoma
		;Gui 1:Add, Button, x320 y18 w45 h18, DOWN
		;Gui 1:Font, S20 cD3FC00, ariel
		;Gui 1:Add, Text, x380 y5, %CurrentWATitle%
		;Gui 1:Add, Text, x380 y5, %A_Space%                                                                                                                    . ;clear the bar
		SavedWATitle:=CurrentWATitle
		GOSUB, RUN_ME_TO_SETUP_VOLUME_UP
	}
}
return

RUN_ME_TO_SETUP_VOLUME_UP:
{
	DetectHiddenWindows, On
	Target := "ahk_class Winamp v1.x"
	IfWinExist %Target%
	{
		SendMessage, 1024, 0, 104   ; IsWinampPlaying?
		WinampPlayStatus := ErrorLevel
		If( WinampPlayStatus == 1 )

		X = 0
		Loop 
		{
			x += 1
			{
			ControlSend, ahk_parent, {up}  ; volume up
			ControlSend, ahk_parent, {up}
			}
			
			if (x > 100)
			{
				SoundBeep
				break
			}
		}
	}
}
return

