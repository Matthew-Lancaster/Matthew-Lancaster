#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;--------------------
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 09-FLICKR KEY RIGHT ARROW __ TIMER __  UNTIL MOUSE RIGHT BUTTON _TOGGLE.ahk
;--------------------

;SHARE URL LINKER GOOGLE DRIVE
;-------------------------------------------------------------
;https://drive.google.com/open?id=0BwoB_cPOibCPLUtONHVlQWNPRDA
;----
;Autokey -- 09-FLICKR KEY RIGHT ARROW __ TIMER __ UNTIL MOUSE RIGHT BUTTON _TOGGLE.ahk - Google Drive
;https://drive.google.com/file/d/0BwoB_cPOibCPLUtONHVlQWNPRDA/view
;----
;-------------------------------------------------------------
;ONLY IN FILE DOWNLOAD UNLESS YOUR GOOGLE DRIVE HAS A ALL DOCUMENT VIEWER LIKE PROGRAM CODE 
;TYPE OF THING
;-------------------------------------------------------------


;--------------------
#SingleInstance force
;--------------------

;---------------------------------------------------------
; CODE INITIALISE SOUND EFFECT LEARN
;---------------------------------------------------------
SoundBeep , 1500 , 400
;---------------------------------------------------------

; --------------------------------------------------------------------------------------------------------
; FOR LOGITECH SEND KEY INPUT MAKE THE CAP LOCK MESSAGE FLASH 
; REQUIRE TO TURN OFF CAP LOCK OR NUM LOCK MAYBE
; THIS CODE DO IT FOR YOU
; --------------------------------------------------------------------------------------------------------
SetCapsLockState ,Off
SetNumLockState , Off
SetScrollLockState , Off


;--------------------------------------------------------------------------------------------
; THE CODE A BIT SKETCHY UNABLE TO USE START STOP
; TRY IN ANOTHER CODE BEFORE NOT SUCCESS BUT AND HERE
; AND NOW LEARN DO TIMER TO SUB ROUTINE WHICH IS CLEARER
;--------------------------------------------------------------------------------------------
; Wed 01 February 2017 06:15:40----------
;--------------------------------------------------------------------------------------------

#Persistent

;----------------------------------------------------------------------
; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
;----------------------------------------------------------------------
; GOES WITH THE ROUTINE AT END NORM IS THEY LIVER NEXT EACH OTHER BUT AND I SPLIT THEM
; THIS PART HAS TO RUN AT ENTRY
;----------------------------------------------------------------------

;DetectHiddenWindows, On
;DetectHiddenWindows, On

Saved_CHROME_Title=
x2 = 0

;--------------------------------------------------------------------------------------------------------------------
; CALL SUB ROUTINE __ Flickr_SUB __ EVERY 1000 MILLISECOND 1 SECOND
;--------------------------------------------------------------------------------------------------------------------
; 400 MS IS OFTEN NICE ON DESKTOP BUT AND PICTURE IMAGE BUILDER TIME INTERNET ON-LINE
;--------------------------------------------------------------------------------------------------------------------
; WHEN COMPUTER LESS QUICKER SEEM ERROR MORE TOWARD ENDER AS LOAD BEEN THROUGH
; ADJUST THE SPEED LESS QUICK
; GUESS EXTRA KEY ARE IGNORED
;--------------------------------------------------------------------------------------------------------------------
;--------------------------------------------------------------------------------------------------------------------
setTimer Flickr_SUB,1000
setTimer Flickr_KEEP_FOCUS_ON_SUB,1000
;--------------------------------------------------------------------------------------------------------------------
setTimer RIGHT_BUTTON_PAUSE_SUB,1
setTimer LEFT_AND_RIGHT_BUTTON_PRESS_TO_EXIT_SUB,10
;--------------------------------------------------------------------------------------------------------------------

return

;----------------------------------------------------------------------
Flickr_SUB:
	IfWinNotActive ahk_class Chrome_WidgetWin_1
	{
		; PAUSED IS X2 = 1
		if x2 = 1
		{
			; IF FOCUS MOVE FROM CHROME AND PAUSED STATE 
			; DON'T DO COMMAND TO REGAIN WINDOW ACTIVE ONTO CHROME
			; NON OPERATION
		}
		else
		{
			winactivate, ahk_exe chrome.exe
			winwaitactive, ahk_exe chrome.exe
			SoundBeep , 1500 , 200
			x = 1
		}
	}
	
	IfWinActive ahk_class Chrome_WidgetWin_1
	{
		;------------------------------------------------------------------------------------------------------------
		; HERE WILL DETECT CHROME IS ACTIVE AND ALSO HELP IN SET IT FOCUS ON
		; IF UNDER THE CURSOR 
		; ABLE TO TELL BY MINIMISE THE WINDOW IN FRONT AND SEE-ER
		
		; SOMETIME REQUIRE A NUDGE
		; AND SET CODE TO BRING FOCUS TO THE CHROME BEEN DONE
		; AND SET TO REMOVE EXIT WHEN NOT IN FOCUS AFTER FOUND HAS BEEN DO
		; WITH BEEP ON
		;------------------------------------------------------------------------------------------------------------
		; ADDITIONAL CODE TODAY 20 FEB
		; AND NOTICE NUMBER BUT NUMBER 1 FIRST
		; THE CODE NOW WILL ONLY FLICK ON TO THE NEXT PICTURE WHEN THE TITLE HAS CHANGED
		; AVOID REPEAT KEY PRESS
		; UPON WORKER NOTICE SECOND THING
		; IF HOLD CURSOR OVER THE LEFT OR RIGHT OR MAINLY RIGHT SIDE OF PICTURE WHERE WOULD PUSH FLIP TO NEXT PICTURE
		; IT SEEM TO PUSH OVER THE WORKING OPERATION ANYWHERE ELSE THE CODE OPERATE GOOD
		; AND ALSO WITH IMAGE FLIPPER GOOD KE PRESS ONLY REQUIRED ADD A SOUND EFFECT FOR PAGE FLIPPER
		;------------------------------------------------------------------------------------------------------------
		; ALSO WHEN AT LAST PICTURE IT WILL STAY THERE
		;------------------------------------------------------------------------------------------------------------
		WinGetTitle Current_CHROME_Title,ahk_class Chrome_WidgetWin_1

		if Current_CHROME_Title<>%Saved_CHROME_Title%
		{
			Saved_CHROME_Title:=Current_CHROME_Title
			winwaitactive, ahk_exe chrome.exe
			;Sleep, 10

			if x2 = 0 
			{
			
				;--------------------------------------------------------------------------------
				; BETTER VERSION TO USE MORE SPEEDY AND ANOTHER
				;--------------------------------------------------------------------------------
				; LENGHT OF TONE ADJUSTMENT
				SoundBeep , 1800 , 20
				;SoundBeep , 1800 , 10
				;SoundBeep , 1800 , 5
				
				Sendinput {Right}

				;--------------------------------------------------------------------------------
				; OR WHORE
				;--------------------------------------------------------------------------------
				;Send, {Right}
				
			}
		}
	}
return
;----------------------------------------------------------------------

;----------------------------------------------------------------------
Flickr_KEEP_FOCUS_ON_SUB:

	;------------------------------------------------------------
	; KEEP FOCUS ON TO MAKE SURE
	; OR REQUIRE GIVE A NUDGE WITH YOUR OWN MANUAL MOUSE CLICKER TO MOVE FORWARD ON
	; ESPECIALLY WHEN TESTER THE PAUSE STATE OFF AND ON
	;------------------------------------------------------------
	IfWinNotActive ahk_class Chrome_WidgetWin_1
	{
		; PAUSED IS X2 = 1
		if x2 = 1
		{
			; REPEAT REMMER BLOCK FROM ABOVE
			; IF FOCUS MOVE FROM CHROME AND PAUSED STATE 
			; DON'T DO COMMAND TO REGAIN WINDOW ACTIVE ONTO CHROME
			; NON OPERATION
		}
		else
		{
			winactivate, ahk_exe chrome.exe
			winwaitactive, ahk_exe chrome.exe
			x = 1
		}
	}
return

;----------------------------------------------------------------------

;----------------------------------------------------------------------
RIGHT_BUTTON_PAUSE_SUB:
;----------------------
; MOUSE RIGHT BUTTON __ P FOR PRESS 
; TO DO WITHER PAUSE
; ----------------------------------------------------------------------
if GetKeyState("RButton", "P") 
{

	;loop, 50
	{
		;SendInput {Esc}   ;____ Kill the context menu
		;sleep 1
	}
	KeyWait, Rbutton ;____ As soon as MOUSE RButton is released...
	;------------------------------------------------------------------------------------------
	; REM STATEMENT SAME NOTE ABOUT __ THING BELOW ___ NEAR
	;------------------------------------------------------------------------------------------
	loop, 20
	{
		SendInput {Esc}   ;____ Kill the context menu
		sleep 1
	}

	IF x2 = 1
	{
		; UNPAUSED
		x2 = 0 
		Saved_CHROME_Title=
		; DOUBLE BLEEPER TO SHOW UNPAUSED
		SoundBeep , 1500 , 400
		sleep 100
		SoundBeep , 2000 , 400
	return
	}

	IF x2 = 0
	{
		; PAUSED
		x2 = 1
		Saved_CHROME_Title=
		; DOUBLE BLEEPER TO SHOW PAUSED
		SoundBeep , 2000 , 400
		sleep 100
		SoundBeep , 1500 , 400
	
	return
	}
}
; ----------------------------------------------------------------------
	
	
;----------------------------------------------------------------------
LEFT_AND_RIGHT_BUTTON_PRESS_TO_EXIT_SUB:	
	;----------------------
	; MOUSE LEFT_ BUTTON __ P FOR PRESS 
	; MOUSE RIGHT BUTTON __ P FOR PRESS 
	;----------------------

	; -----------------------------------------------------------------------------------------------------------
	; THE CONTEXT MENU UPON RIGHT CLICK MOUSE BUTTON
	; DOES NOT ALWAYS DISAPPEAR TO THE ESCAPE KEY PRESSOR
	; AND ANOTHER SOLUTION ADDITIONAL
	; DON'T WAIT FOR THE MUSE RIGHT CLICKER BUTTON TO RELEASE BACK UP
	; t is better not TO PUT AND ESC KEY BEFORE MOUSE RIGHT BUTTON CONTEXT MENU IS RELEASED
	; AND SUCCESSFULLY LOOP A FEW TIME IS ABLE MAKE DISAPPEAR CONTEXT MENU
	; BUT TOO MANY LOOP WITH SLEEP AND HOLD UP RESPONSIVENESS
	;
	; ONLY THE CONTEXT MENU IS LITTLE FAULTY TO HIDE
	; DATE TIME END PROJECT AT NEW LEVEL
	; Thu 02 February 2017 01:59:30----------
	
	; -----------------------------------------------------------------------------------------------------------
	if GetKeyState("RButton", "P") and GetKeyState("LButton", "P") 
	{
	KeyWait, Rbutton ;____ As soon as MOUSE RButton is released...
	loop, 20
	{
		SendInput {Esc}   ;____ Kill the context menu
		sleep 1
	}

	SoundBeep , 2000 , 400
	sleep 100
	SoundBeep , 2200 , 400

 	Exitapp
	}
	
	IfWinNotActive ahk_class Chrome_WidgetWin_1
	{
		; --------------------------------------------------------------------------------------------------------------------
		; NON OPERATION REMMER ____
		; IF PAUSED STATE HAPPEN AND NOT IN FOCUS GET A REPEAT BEEPER FROM HERE
		; --------------------------------------------------------------------------------------------------------------------
		; SoundBeep , 1800 , 400

		If x=1
		{
		;SoundBeep , 1800 , 400
		;Exitapp
		}
	}
return
;----------------------------------------------------------------------

;----------------------------------------------------------------------
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
;----------------------------------------------------------------------