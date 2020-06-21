;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk
;# __ 
;# __ Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;--------------------
; Autokey -- 10-Read Mouse Cursor Icon State And Beeper When Not Busy Hour Glass Over
;--------------------
;  :\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk
;--------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; ONLINE LOCATE HERE
; -------------------------------------------------------------------
; IS THE WHOLE FOLDER OF AUTOHOTKEYS IN MY CODE
; THIS IS THE FILE SHARE BUT NOT SURE IF STAY WITH THE LINK WHEN UPDATES OCCUR
; SO I PUT THE WHOLE FOLDER
; LOOK FOR THE TITLE ABOVE SCRIPT NUMBER 10
; ALSO THERE IS ANOTHER FOLDER FOR THE SOUND EFFECT I USE HERE TO MAKE IT UP
; OR YOU CAN DO THE REM LINE AND GO BACK TO SOUND BEEPER
; IN THE FOLDER \AUDIO RECORDER MOUSE
; THEY ARE SOME TONES OF THE DTMF AND PLAYED ABOUT WITH AND REDUCED VOLUME 
; SO AUTOHOTKEYS VOLUME MIXER CAN BE LEFT ON FULL VOLUME WITH THIS SOUNDING TOO LOUD
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
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
; https://www.dropbox.com/s/puf965euzqeslkp/Autokey%20--%2010-READ%20MOUSE%20CURSOR%20ICON%20STATE%20AND%20BEEPER%20WHEN%20NOT%20BUSY%20HOUR%20GLASS%20OVER.ahk?dl=0
;# ------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------


;----------------------------------
; System Programmer
;----------------------------------

; ---------------------------------------------------------------
; CODE WANTER FOR REASON GOODSYNC IS NEAR IMPOSABLE TO GET GOING FOREVER
; I have BEEN WRITING THE OWNER FOR NIGHTMARE OVER 2 YEAR
; AND MUCK ABOUT NOT WANT TO GIVE ANYTHING TOWARD MAKE SORT CONFLICT OUT
; LIKE LATEST REQUEST IS SAME AS FIRST
; ONE WAS TO EXPAND ALL THE-NODE TREE BRANCH TO SEE CONFLICT EVERYWHERE
; OR BUNCHED AWAY HIDDEN

; NEXT THE SECOND ON
; WOULD BE IF NOT NODE EXPAND TREE
; THEN JUMP TO NEXT CONFLICT

; AND OTHER PROBLEM
; WHILE DISCOVER MANY AND GIVE AWAY THE WORK WHILE ASK OR SOMETHING AND MUCKED ABOUT NOT AN ANSWER
; ASKING FOR A FULL EXAMPLE HOW TO SCRIPT COMMAND GO TOGETHER
; BECAUSE MY TEST THEY BOT WORKER
; I WANTED DATE RESTRICTION AND ALSO AT SUB-FOLDER RATHER THAN ONE LEVEL ROOT PATH
; IF THEY GAVE A WHOLE STRING ONE COMMAND ASSEMBLE 
; THEN I WOULD TEST AND SHOW IS MAYBE WRING

; DO THEY NON NOT THEY DON'T
; ONE OF HALF A COMMAND OR THE OTHER WITH A PUT-DOWN
; ABOUT TEACHER
; COURTROOM ROOM ADVISOR MIDDLE MAN TO THE JUDGE ABOUT THE DATA OWNERSHIP AND MEDIA ELECTRONIC OWNERSHIP UNIVERSITY TYPE

; AND THE SISTER BROTHER PROJECT IS ROBOFORM 
; LIKE LIKE LIKE ROBOT RENTAL WITHER AND CASH ROLLING IN
; HE SYNC IN MINI VERSION ARRIVED AT GOODSYNC

; BUT GOODSYNC ALTHOUGH NOT THE PARTNER IS NOT THE RENTAL SISTER ROBOFORM ARE
; AND LESS CASH FLOW GO TOWARD HER

; MANY SYNC PROGRAM CODE ARE ABLE to USE NEW STYLE OF LOCATE A WHOLE FOLDER MOVED
; BUT NONE ARE BETTER CHECK WIKIPEDIA
; AND I BUY A FEW WHICH ARE LINED UP WITH ALL STAR RATING OF FUNCTION
; BUT NOT NEAR-ER AS WELL EQUIPPED
; TRY THEM OTHER OUT AND GET MONEY ON REFUND

; RENT A ROLL OF CASH ROBOFORM

; ALSO I DON'T LIKE IT 
; THE DATA FOR PROJECT IS HARD WORK SYSTEM DATABASE TO EXTRACT USE IN MANY WORK PROJECT
; BUT ND ALSO HE HAS SYSTEM TO SAVE IN MORE READABLE FORMAT
; BUT HE WON'T HELP ME CHANGED IT OVER BY THE MODIFY THE REGISTRY REQUIRED TALK ABOUT IN MANUAL

; AND LAST NOTE 
; I NOW USE THE HUBIX.COM CLOUD
; AND THAT THING IS ABLE REMOTELY TO DO THE WHOLE BACKUP
; WITH ANY SYNC PROBLEM 
; AND NOT BEEN ANY BIG MEMORY USER

; STILL GOING TO BE  MORE PROBLEM LATER IF I EVENTUALLY GET IT RUNNER

; ---------------------------------------------------------------
; [ Sunday 17:53:10 Pm_15 April 2018 ]
; WORK TO REMOVE AUDIO CLICKS WHEN ONE SOUND FILE IS PLAYED OVER ANOTHER TOO QUICKLY _ TIMER DELAY ANSWER 
; ---------------------------------------------------------------
; PREVIOUSLY TO DATE ABOVE
; I ADDED A CODE THAT DETECT IF THE AUTOHOTKEYS SCRIPT IS ALREADY RUNNING AND THEN EXIT ITSELF
; BECAUSE ON LESS QUICK MACHINE WHEN I LOAD MY ALL 8 AHK'S SCRIPT SOME DOUBLE UP BECAUSE THEY ARE LOADED SO QUICKLY IT SUPPOSED TO HAVE A SAFEGUARD THAT IT DOESN'T HAPPEN
; ALL MY AHK SCRIPT'S INCLUDE THAT THING
; ---------------------------------------------------------------
; THERE IS AN AWFUL LOT OF IRRELEVANT REM STATEMENT INCLUDE IN THIS CODE AS I HAD A BIG PLAY ABOUT AND NEVER TOOK MOST THEM OUT
; ---------------------------------------------------------------

; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
; USE HERE #Persistent OR WILL LEAVE THE CODE UPON ONE CYCLE OF A RUNNER
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; ---------------------------------------------------------------
; I MADE MENU ITEM INTO INCLUDE FILE IN 3 PART 
; 01. INTRO SETUP MENU
; 02. THE MENU ROUTINE
; 03. ANY ROUTINE THE MENU USE
; ---------------------------------------------------------------
; SAVER OF RSI INJURY AND MORE ACCURATE
; THE INCLUDE FILE ARE SAME FOLDER
; ---------------------------------------------------------------
; FROM __ Sun 09-Jun-2019 07:03:00 __ Clipboard Count = 024
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------
; Create the popup menu by adding some items to it.
; MenuHandler:
; ---------------------------------------------------------------
; #Include GO WITH FULL PATH AS SOME LAUNCHER DO NOT SET WORK PATH WHEN RUNNER
; RATHER THAN CHANGE THE WORKING PATH WITHIN-AH
; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk


;---------------------------------------------------------
; CODE INITIALIZE SOUND EFFECT LEARN
;---------------------------------------------------------
SoundBeep , 1500 , 400
;SoundSetWaveVolume, 100 
;SoundSet, 5
;---------------------------------------------------------


; ---------------------------------------------------------------
; THE CODE A BIT SKETCHY UNABLE TO USE START STOP
; TRY IN ANOTHER CODE BEFORE NOT SUCCESS BUT AND HERE
; AND LEARN DO TIMER TO SUB-ROUTINE WHICH IS CLEARER
; ---------------------------------------------------------------
; Wed 01 February 2017 06:15:40----------
; ---------------------------------------------------------------


DetectHiddenWindows, On

Saved_MOUSE_CURSOR_Title = %A_Cursor%
RELEASE_SOUNDPLAY=%A_Now%
ALLOW_SOUND=1


; NOT REQUIRE HERE IS NOT AUTO RUN FOR THE COMPUTER THAT ARE OTHER ONE
; SET_GO=FALSE
; IF (A_ComputerName = "4-ASUS-GL522VW") 
; SET_GO=TRUE
; IF (A_ComputerName = "7-ASUS-GL522VW") 
; SET_GO=TRUE
; IF (A_ComputerName = "8-MSI-GP62M-7RD") 
; SET_GO=TRUE

; IF SET_GO=FALSE
	; PAUSE


; -------------------------------------------------------------------
; TEST RUN IN CODE STARTUP
; FILE NAME TOO LONG HERE AND WON'T LIKE IT NEVER PLAY
; -------------------------------------------------------------------
; DON'T WANT STARTUP SOUND ON MY MAIN COMPUTER
; USEFUL IF YOUR EVER GOING TO LEARN WHAT TO GET
; -------------------------------------------------------------------

IF !(A_ComputerName = "7-ASUS-GL522VW") 
	Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\AutoHotKeys Mouse Changer _ Wait _ Hour Glass.wav



; ---------------------------------------------------------------
; CALL SUB-ROUTINE __ MOUSE_CURSOR_SUB __ EVERY 1000 MILLISECOND 1 SECOND
; ---------------------------------------------------------------
; ---------------------------------------------------------------
setTimer MOUSE_CURSOR_SUB,1
setTimer TIMER_PREVIOUS_INSTANCE,1
setTimer RELEASE_SOUNDPLAY_TIMER,1000


return

; ---------------------------------------------------------------
; ---------------------------------------------------------------


RELEASE_SOUNDPLAY_TIMER:
	If (A_Now>=RELEASE_SOUNDPLAY)
	{
		; RELEASE THE SOUND WAV FILE BEING LOCKED HOLDING FOR SYNCHRONISE
		SoundPlay, Nonexistent.avi
		setTimer RELEASE_SOUNDPLAY_TIMER, OFF
	}
RETURN

;--------------------------------------------------------------------
MOUSE_CURSOR_SUB:

Current_MOUSE_CURSOR_Title = %A_Cursor%

SOUND_PLAYED=0

; ---------------------------------------------------------------
; SET THE TIMER QUICK WHEN NOTHING IS HAPPEN FOR READY EVENT
; OR WHEN THE SOUND PLAY HAS A SLIGHT WAIT TIMER FOR EVENT SOUND DONE
; ---------------------------------------------------------------

;if %Current_MOUSE_CURSOR_Title%=%Saved_MOUSE_CURSOR_Title%
;{
	;TIMER_DURATION_VAR=20
	;setTimer MOUSE_CURSOR_SUB,10
;}


 
SET_GO=0
if Current_MOUSE_CURSOR_Title<>%Saved_MOUSE_CURSOR_Title%
	SET_GO=1

ALLOW_SOUND=1
IF SET_GO=1
{
	isFullScreen := isWindowFullScreen( "A" ) ; ActiveWindow
	if isFullScreen 
		ALLOW_SOUND=0
}

FN_NAME:="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 24-I_VIEW32 CONVERT_CCSE.AHK"
IF WinExist(FN_NAME) 
	ALLOW_SOUND=0

; -------------------------------------------------------------------
; THIS IS WHAT I ORIGINALLY MADE THIS THING FOR
; TO INDICATE CURSOR CHANGE AFTER LONG HOUR GLASS SO I NOT FALL ASLEEP WAIT 
; AT SCREEN FO LONG SYBC-ING WITH GOODSYNC
; NOW THE NOISE ANNOY ME A BIT ON MY MAIN COMPUTER UPSTAIRS IN BED
; SO RESTRICTING IT A BIT HERE ONLY FOR GOODSYNC ON THIS COMPUTER
; IT WAS ALSO TO FIND LAND ON THE CURTAIN MIDDLE FRAME BA LEFT AND RIGHT MOVE
; WHICH IS SO HARD TO LAND ON A THIN LINE
; THIS DOES IT WITH ECHO SOUNDING
; MUCH MORE HELPFUL AND GOOD IN OTHER PROGRAM WHEN WANT IT ON
; -------------------------------------------------------------------
IF (A_ComputerName = "7-ASUS-GL522VW") 
	IfWinNotActive, GoodSync -
		ALLOW_SOUND=0
; -------------------------------------------------------------------

ALLOW_SOUND=0
IfWinActive, GoodSync -
	ALLOW_SOUND=1
IfWinActive, GoodSync2Go -
	ALLOW_SOUND=1
	
IF ALLOW_SOUND=0
	SET_GO=0
	
if SET_GO=1
{
	Saved_MOUSE_CURSOR_Title = %Current_MOUSE_CURSOR_Title%
	SOUND_PLAY_TRUE=0
	
	RELEASE_SOUNDPLAY=%A_Now%
	RELEASE_SOUNDPLAY+=4 ; 4 SECONDS
	setTimer RELEASE_SOUNDPLAY_TIMER,1000

	
	; ---------------------------------------------------------------
	; IF Saved_MOUSE_CURSOR_Title=  ____ HAS _NOTHING_ FIRST TIME RUN IN __ GIVE IT THE FIRST VALUE OF CURRENT
	; ---------------------------------------------------------------
	; IF ONE COMMAND AFTER IF CONDITION  AND THEN NOT REQUIRE 
	; A BRACKET PAIR AROUND _LIKE LIKE LIKE_ {... _NEW LINE_VBCRLF_...} A SET 
	; OF LINE STATEMENT COMMAND
	; ---------------------------------------------------------------
	; LATER SAME COMMAND BE USER ANOTHER FOR MAIN LOOP ALWAYS HAPPEN LATER ALTERNATIVE AFTER ONCE
	; ---------------------------------------------------------------
	; if Saved_MOUSE_CURSOR_Title=
	; Saved_MOUSE_CURSOR_Title = %Current_MOUSE_CURSOR_Title%

	; ---------------------------------------------------------------
	; ***********************
	; SET SOUND PLAY ONCE WHATEVER EVENT MIGHT OCCUR
	; NOT GOOD TO SET BEFORE THE IF CONDITION HERE ONLY WITHINER
	; ---------------------------------------------------------------
	; SOUND_PLAY_TRUE=0

	; ---------------------------------------------------------------
	; DEBUG TEST HERE MORE COMMON THAT WAIT HOUR GLASS
	; WORKER HAPPEN 
	; ---------------------------------------------------------------
	;if Current_MOUSE_CURSOR_Title=IBeam
	;SoundBeep , 2500 , 4000

	; ---------------------------------------------------------------
	; IF MOUSE CURSOR WAS WAIT HOUR GLASS AND RETURN TO RESULT WANTED THEN HERE
	; ---------------------------------------------------------------
	;---------------------------------------------------
	; if Saved_MOUSE_CURSOR_Title=IBeam __ TEST DEBUGGER REMMER OUT
	; QUICKER RESULT TO TEST THAN MAKE THE HOUR GLASS APPEAR
	; --------------------------------------------------
	
	
	if Saved_MOUSE_CURSOR_Title=Wait
	; ---------------------------------------------------------------
	; TONE TONE SOUND AUDIO HERALD THE SOUND FOR CHANGE WANTER HAPPENING
	; ---------------------------------------------------------------
	
	; ---------------------------------------------------------------
	{
		
		; -----------------------------------------------------------
		; return ____
		; HERE AND EXIT CONTINUE FURTHER SO NOT PLAY A SOUND DOUBLE FOR THE OTHER _ ALL CHANGE OF CURSOR
		; OR SET A VARIABLE DO
		; OR WHORE MIGHT BE BETTER HAVE VARIABLE NOT PLAY OTHER SOUND FOR CODE SAKER
		; AS WHEN GET THERE WILL SHOW DEBUGGER
		; SMALL CODE
		; -----------------------------------------------------------
		; return ____
		; DON'T WORK AS EASY THAT MEHTODIER ABOVE THROW ERROR SOUND REPEAT WITH NONE CHANGE
		; ERROR IN CODE CHECK FOR _0 AT END NOT _1 OPPOSITE
		; -----------------------------------------------------------
		; MOVE ON TO NEXT METHOD STEP 2 VARIABLE
		; BETTER VARIABLE INSTEAD
		; -----------------------------------------------------------
		
		; Wait __ HOUR GLASS
		
		;SoundBeep , 1800 , 150
                
		;SLEEP 150
		;TIMER_DURATION_VAR=500
		;SoundBeep , 2000 , TIMER_DURATION_VAR
		
		;SoundBeep , 2000 , 400

		;###
		;### Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\start_VOID_EMPTY_FILE_STOP_SOUND_PLAY.wav
		;-----------------------------------

		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\AutoHotKeys Mouse Changer _ Wait _ Hour Glass.wav
		;-----------------------------------
		
		SOUND_PLAY_TRUE=1
		SOUND_PLAYED=1

		;setTimer,, 500
	}
	
	
	SET_GO=0
	if Saved_MOUSE_CURSOR_Title=AppStarting
		SET_GO=1
	if SOUND_PLAY_TRUE=1
		SET_GO=0
	IF SET_GO=1 
	
	; ---------------------------------------------------------------
	; ARROW AND HOURGLASS FOR APP STARTING
	; TONE TONE SOUND AUDIO HERALD THE SOUND FOR CHANGE WANTER HAPPENING
	; ---------------------------------------------------------------
	{
		; -----------------------------------------------------------
		; SHORTER BLIPER FOR LOAD APPLICATION SYMBOL MOUSE CURSOR IS ARROW AND HOUR GLASS
		; HOUR OUCH HURT TICKER __ GLASS METAL SAND WATER FLAME __
		; -----------------------------------------------------------
		
		; AppStarting

		;SoundBeep , 2000 , 100

		;TIMER_DURATION_VAR=500
		;setTimer,, 500

		;SoundBeep , 1800 , 400

		;### Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\start_VOID.wav
		;-----------------------------------
		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\AutoHotKeys Mouse Changer _ App Starting.wav

		SOUND_PLAY_TRUE=1
		SOUND_PLAYED=1
	}

	;Saved_MOUSE_CURSOR_Title:=Current_MOUSE_CURSOR_Title

	; ---------------------------------------------------------------
	; CARE ABOUT THE SOUND NOT TOO LONG, 200 MILLISECOND IS ENOUGH
	; OR SYSTEM DRAG LESS QUICKER
	; ---------------------------------------------------------------
	; USE BEEP FOR EVERY CHANGE OR REM OUT	
	; OR VARIABLE BLOCK OUT UNLESS CHANGE NOT HAPPEN ELSEWHERE
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	; SOME RETURN FROM STATE AS A EXTRA BLEEPER ARRIVE WITHOUT SHOWN REASON
	; DEBUGGER
	; ---------------------------------------------------------------
	
	; ---------------------------------------------------------------
	;DEBUG_SOUND_PLAY_TRUE=1
	

	;if Saved_MOUSE_CURSOR_Title=IBeam -----------------
	;DEBUG_SOUND_PLAY_TRUE=0 
	;---------------------------------------------------
	; IBEAM HAS A PROBLEM OF REPEAT
	;---------------------------------------------------
	
	;if Saved_MOUSE_CURSOR_Title=AppStarting 
	;DEBUG_SOUND_PLAY_TRUE=0 
	
	;if Saved_MOUSE_CURSOR_Title=Wait 
	;DEBUG_SOUND_PLAY_TRUE=0 
	
	;if DEBUG_SOUND_PLAY_TRUE=1
	;{
	;	ListVars
    ;    Pause
	;}
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	; PROBLEM IF RESPONSE NOT HAPPEN VERY WELL MIGHT BE TIME FOR REBOOT
	; ---------------------------------------------------------------
	; REQUIRE TIMER VAR NOT TO PLAY SOUND STACKER WHEN ONE QU-ED UP AND 
	; PLAYING FOR DURATION MICRO SECOND
	; SEEM OKAY AT REBOOT BUT LAGG LATER
	; ---------------------------------------------------------------

	if SOUND_PLAY_TRUE=0
        {

		;TIMER_DURATION_VAR=40
		;SoundBeep , 5000 , TIMER_DURATION_VAR
		
		;SoundBeep , 2000 , 400
		;### Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\start_VOID.wav

		;-----------------------------------
		Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\AutoHotKeys Mouse Changer Normal.wav
		SOUND_PLAYED=1
		
		;SoundBeep , 3000 , 40

		;setTimer,, 20
        }
		
	; ---------------------------------------------------------------
	; EXAMPLE FOR REASON PROJECT __ EXTRA LEARN SOUNDER ALARMER
	; FIND WHEN MOUSE GIVES LEARN RETURN FROM HOUR GLASS WAIT STATE
	; ---------------------------------------------------------------
	; I USE TO PLAY WITH MOUSE CURSOR CODE IN VB 6 BACK IN THE DAY QUITE A LOT 
	; AND GET LOST ON THAT ONE
	; REASON TO CHANGE IT
	; MOVE AROUND THE SCREEN ANYWHERE WANT THAT EASY PART
	; FIND UNDER POINT X-Y POINT COMMAND EASY ANOTHER
	; TAKE TIME LEARN THEM ALL
	; THAT MAKE IT EASIER FOR ME TO FIND THE SEARCH TERM CONSTRUCTION 
	; TO AUTOHOTKEYS
	; ---------------------------------------------------------------
	
	;TESTER = IBeam
	;TESTER = Wait
	;StringGetPos, Pos_FIND_MOUSE_SHAPE_IS_PAUSE_WAIT_STATE__1, Current_MOUSE_CURSOR_Title, %TESTER%
	;if Pos_FIND_MOUSE_SHAPE_IS_PAUSE_WAIT_STATE__1 > 0 
	;SoundBeep , 2000 , 2000
	; ---------------------------------------------------------------
	; ABOVE LESS SKILL METHOD FOR THE PROJECT __ MORE OVER KILL OTHER WAY AROUND WITH COMPLEX
	; ---------------------------------------------------------------

	;if Current_MOUSE_CURSOR_Title<>%Saved_MOUSE_CURSOR_Title%
	;Saved_MOUSE_CURSOR_Title:=Current_MOUSE_CURSOR_Title
	
	;--------------------------------
	; DEBUG TEST RESULT
	;--------------------------------
	; MSGBOX, %Current_MOUSE_CURSOR_Title%
	;--------------------------------
	
	;TESTER = IBeam
	;StringGetPos, Pos_FIND_MOUSE_SHAPE_IS_PAUSE_WAIT_STATE__1, Current_MOUSE_CURSOR_Title, %TESTER%
	;if Pos_FIND_MOUSE_SHAPE_IS_PAUSE_WAIT_STATE__1 > 0 
	;SoundBeep , 2000 , 2000
	
	;----------------------------------------------------------------
	; GET CLICKS ON SOUND AUDIO WHEN QUICKER AND NOT SLIGHT PAUSE BETWEEN SOUND PLAYED
	;----------------------------------------------------------------
	if SOUND_PLAYED=1
		setTimer MOUSE_CURSOR_SUB,100
	if SOUND_PLAYED=0
	setTimer MOUSE_CURSOR_SUB,1
	
}

return
;--------------------------------------------------------------------
		


isWindowFullScreen( winTitle ) {
 ; checks if the specified window is full screen
 
  winID := WinExist( winTitle )

  If ( !winID )
  Return false

  WinGet style, Style, ahk_id %WinID%
 WinGetPos ,,,winW,winH, %winTitle%
 ; 0x800000 is WS_BORDER.
 ; 0x20000000 is WS_MINIMIZE.
 ; no border and not minimized
 Return ((style & 0x20800000) or winH < A_ScreenHeight or winW < A_ScreenWidth) ? false : true
;# ----
;# Detect Fullscreen application? - Ask for Help - AutoHotkey Community
;# https://autohotkey.com/board/topic/38882-detect-fullscreen-application/
;# ----
}





#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk

	
		
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
	return (match > 1)
	}
Return


		
;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------
		
; ---------------------------------------------------------------
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
; ---------------------------------------------------------------


;---------------------------		
; END OF CODE REMMER INFO
;---------------------------

; LINK FOR HELPING SOURCE __ LEAD 02 OF 02
;---------------------
;Sun 05-Feb-2017 02:54:50
;---------------------
; Form FindWindow ---
; Variables and Expressions - Google Chrome
;---------------------
;----
; Variables and Expressions
;https://www.autohotkey.com/docs/Variables.htm#Cursor
;----
;---------------------
;---------------------
; Form FindWindow ---
; Variables and Expressions - Google Chrome
;---------------------
; Misc.

;A_Cursor	
; The type of mouse cursor currently being displayed. It will be one of the following words: AppStarting, Arrow, Cross, Help, IBeam, Icon, No, Size, SizeAll, SizeNESW, SizeNS, SizeNWSE, SizeWE, UpArrow, Wait, Unknown. The acronyms used with the size-type cursors are compass directions, e.g. NESW = NorthEast+SouthWest. The hand-shaped cursors (pointing and grabbing) are classified as Unknown.
;---------------------

; ---------------------------------------------------------------
; LINK FOR HELPING SOURCE __ PROVIDE 01 OF 02
; ---------------------------------------------------------------
;---------------------
;Sun 05-Feb-2017 02:56:50
;---------------------
; Form FindWindow ---
; How to get the mouse cursor's shape? - Gaming Questions - AutoHotkey Community - Google Chrome
;---------------------
;----
; How to get the mouse cursor's shape? - Gaming Questions - AutoHotkey Community
;https://autohotkey.com/board/topic/101626-how-to-get-the-mouse-cursors-shape/
;----

; ---------------------------------------------------------------
; MY FIRST AND ONLY SEARCH TERM
; ---------------------------------------------------------------
;---------------------
;Sun 05-Feb-2017 02:57:15
;---------------------
; Form FindWindow ---
; AUTOHOTKEYS READ THE MOUSE CURSOR SHAPE - Google Search - Google Chrome
;---------------------
;----
; AUTOHOTKEYS READ THE MOUSE CURSOR SHAPE - Google Search
;https://www.google.co.uk/search?q=AUTOHOTKEYS+READ+THE+MOUSE+CURSOR+SHAPE&rlz=1C1CHBD_en-GBGB721GB721&oq=AUTOHOTKEYS+READ+THE+MOUSE+CURSOR+SHAPE&aqs=chrome..69i57.13982j0j7&sourceid=chrome&ie=UTF-8
;----
;---------------------

; ---------------------------------------------------------------
; EXTRA INFO OF ANOTHER PROJECT ABOUT UNDER THE MOUSE CURSOR
; FOR LINK FINDER IN BROWSER
; ---------------------------------------------------------------
;----
; How to detect mouse cursor state over link? - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/81381-how-to-detect-mouse-cursor-state-over-link/
;----

; Now First Quarter @ 04 Feb 04:20:02
; Next Waxing Gibbous @ 07 Feb 00:28:31
; in 1 Day 20:11:31 H-M-S
; Luminosity Now 58.36609% Day Before 47.76793%

; Uni-Time (UT GMT Solar Atomic) 05-Feb-17 04:17:00
; The UK TimeZone GMT UTC+0 WET

; ---------------------------------------------------------------
; I WILL BE WORKER ALL WEEKEND THIS KIND AGRO THING
; ---------------------------------------------------------------
		
