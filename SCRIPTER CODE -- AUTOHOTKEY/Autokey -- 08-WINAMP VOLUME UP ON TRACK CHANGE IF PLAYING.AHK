#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 08-WINAMP VOLUME UP ON TRACK CHANGE IF PLAYING.AHK

; THIS WILL BE PROJECT 02 OF 02 IN the SAME THEME FOR WINAMP
; AND THIS NUMBER 02 IS THE FAR MORE ADVANCED

; IT IS FOR PROVIDING TO ENHANCE THE FNOOB INTERNET MUSIC STATION
; I HAD TO TURN IT DOWN LAST MONDAY EARLY HOUR
; THEY GONE AT 12 MIDNIGHT SUNDAY
; FOR LITTLE WHILE OFF AIR

; IT IS FULLY COMPLETE AND TESTED EVERY WAY POSSIBLE AS I GOT TIME TODAY
; IT WILL BE FREQUENTLY USE AND IT IS WHAT I WANT

; AND The SECOND CHECK ON ANOTHER COMPUTER AFTER TEXT EDITING OF REMARK FOR SYNTAX ERROR
; AND I DO FIND SOME
; BUT UN-ABLE TO TEST AS THE OTHER SYSTEM IS SET-UP TO RUN WINAMP WITH INTERNET FNOOB TECHNO STATION



;---------------------------------------------------------------------------------------------
;NEW CODE WINAMP CHANGER VOLUME UP ON TRACK CHANGE IF PLAYING ACTIVE
; 17:27 PM TUE 01 NOV
; TO
; 18:27 ENDED
; 19:28 SIGNIFICANT CODE SIGNAL ARRIVED AT - DOCUMENTATION NEARING ENDER FOR DAY
; 19:45 SOURCE FIND CODE RESULT ENTERED
; 19:51 CHECK TEXT FORMATTING NEARING END
; 08:08p -- WEETABIX IS ALL I HAD FOR BREAKFAST -- TIN OF SOUP KNOCKING
; 10:00p SIGNIFICANT LEVEL REACHED OF CHANGE
; 10:20p Another 20 Min to End -- Two Project at Same Time

; 23:17 -- GOING ON-LINE GRAMMAR
; 23:23 -- GOING ON-LINE TYPING -- DONE -- Project 01 of 02
; 23:23 -- GOING ON-LINE TYPING -- DONE -- Project 02 of 02

;Wed 02 November 2016 00:10:21 -- GOING ON-LINE FINAL CHECK OVER -- Project 02 of 02



; WELCOME TO READ THROUGH ALL IS EXPLAINED
; IT IS MY DELIGHT AND IT'S A PLEASURE TO PRESENT THIS CODE 
; AS FREE FOR EDUCATION 
; WORK MY HEAD DOING THIS -- ENJOY THE CODE -- TROUBLE IT TAKEN INTO
; IF I WERE TO SELL MY BLEEDING HEART
; I WOULD TALK ABOUT WHAT THE GOOD THING I WOULD DO WITH MONEY OFFERED BY ANYBODY THAT THOUGHT THIS CODE 
; HELP THEM IN A USEFUL WAY
; I WOULD LIKE TO TELL THE PROJECT I AM DO -- AND HOW ANY CONTRIBUTION TO TECHNOLOGY WOULD HELP
; AND WHAT PROJECT I HAVE
; AND GOAL WITHOUT A GOAL I WOULD NOT ASK 
; TECHNOLOGY AND DATA AS INTELLECTUAL PROPERTY 
; PUT YOUR MARK ON IT -- STINKER
; WHILE I HAVE NOT ANY PROTECTION UP AHEAD
; I WOULD BE GRATEFUL FOR CODE -- IDEA
; RATHER THAN A HOARDING METHOD OF -- MARK IT'S MINE
; I'M NOT EMPLOYER AND  DON'T WANT EMPLOY
; TELL YOU WHAT I COULD DREAM UP SOME FABULOUS HARDWARE
; I DON'T WANT BLEED HEART STORY WHEN SOFTWARE TRIES TO SELL ITSELF
; LIKE LIKE LIKE THE MICROSOFT BILL GATE -- AND OPEN LETTER STORY
; KEEP YOUR SPIRIT UP
; COMPLEX THING ARE QUAD 4 PARALLEL PROCESSOR INTEL 7 MADE IN HONG KONG HK
; COMPLEX FOR CERTAIN ARE SOME SOFTWARE
; LIKE GRUB4DOS -- ALBEIT VERY SMALL BUT IT HELP LOAD THE FRONT END OF WINDOWS
; DESIGNED BY LINUX USER CODER
; THROW OVER TO THE WINDOW COMMUNITY
; ANOTHER COMPLEX -- THE HEART OF SYSTEM PARTITION WIZARD 
; AT LEAST IT MADE POSSIBLE AFTER F-DISK AND WINDOW LIMITED DISK MANAGER 
; PROGRAMMER CODE MADE FOR PROGRAMMING
; WINAMP

; AUTOHOTKEY -- NOT MUCH IMPRESSIVE YET
; IRFANVIEW
; VISUAL BASIC LANGUAGE
; WORK WENT INTO MINI WINDOWS XP AND 7
; MORE TO FOLLOW LIST SCRIPT CREATION

; THE ADVERT BLOCKING SOFTWARE OF CHROME
;TALK ABOUT A CHANGE IN HERE THIS TALK
; CHROME HAS A FEATURE THAT SOME OTHER PEOPLE CAN WRITE WEB PAGE AND TOTAL BLOCK COPY AND PASTE
; BUT FIREFOX STILL ALLOW IT ALL AND BOTH ARE SUPPOSED TO BE A CHARITY SHOP
; BUT GOOGLE PLY TONS IN ADVERT REVENUE AND SECRECY UP THEIR SLEEVE
; THEY TOOK OUT SUPPORT OF WIN XP CHROME
; AND THEN JUST AFTER STOP THE BACKSPACE KEY DESTROY TEXT IN  A FORM WHILE TYPING
; BUT TOO LATE WIN XP
; ALSO ANYTHING MOBILE PAGE IN GRIN BOOK FB FACEBOOK LAND IS NOT ABLE TO COPY PASTE IN CHROME BUT IS ALLOWED IN FIREFOX



; GOING ADVANCED WITH TECHNOLOGY -- PUTTING MORE RESOURCE INTO WHAT IS MODERN

; DON'T GET ME WRONG GOOGLE PROVIDE ME WITH A BLOGGER AND SAVE ME MONEY ON A DOMAIN
; UNLESS I HAD SOMETHING SUCCESSFUL
; I HAD TO GET MY EMAIL FROM ASK AND AMERICAN ON USENET IN GMAIL
; THAT IS WHEN THEIR STUDY TO FIND OUT WHAT IS THE MOST EFFECTIVE
; MY CHOICE OF OFFER CAME FROM 3 AND THE PERSON OFFERED ME I TOOK 
; HAS BLACKLISTED ME FROM RETURN

; MIDNIGHT IS HERE
; GOOD NIGHT MORNING TUESDAY AND WEDNESDAY

; SPELL CHECK IN CAPITAL IS A LITTLE LESS CHECKER
; AND DEPEND ON SERVER LOAD GET A BETTER RESULT
; GRAMMARLY CHECKER


;------------------------------------------------------------------------------------------------------------


;-----------------------------------------------------------------
; PROJECT 01 ---- EXISTING ----------------------------
;-----------------------------------------------------------------
; PROJECT 02 ---- HERE ---------------------------------
;-----------------------------------------------------------------
; PROJECT 02 LEARN SUBROUTINE CALL
; PROJECT 02 LEARN LOOP TO A SUBROUTINE ON TIMER EVENT
;-----------------------------------------------------------------

; IDEA ORIGINATE FROM
;-----------------------------------------------------------------
; VERY NICELY WORKING ON MY LOW ENDER COMPUTER NOW ALL THE PROCESS ARE SORTED UNDER CONTROL
; PITY I NOT ABLE TO LOAD WEB PAGE GOOD ON THIS COMPUTER 
; SO THAT WORK GOES ON ANOTHER MY COMPUTER
; TO FIND RESULT AND LOAD IN SEARCH TERM TO PASTE IT
;
;-----------------------------------------------------------------
; FOR THE END USER OF AUTOHOTKEYS HOOKER HOOKER
;-----------------------------------------------------------------
; YOU WANT TO REALISE THE BENEFIT OF THIS CODE
; THE WORKING IS FROM VISUAL BASIC 6 
; MY VISUAL BASIC HAS A PERIOD SET TIME WHERE IT TURN THE MUSIC OFF
;--- THIS IS ONLY TO TURN IT UP
;--- ALSO MY VISUAL BASIC TURNS OF OFF A MONET WHILE TALKER VOICE SPEECH SPEAKING
;-----------------------------------------------------------------
; THE VOLUME SETTING IS EITHER OFF OR ON
; THE VARIABLE ADJUSTMENT IS YOUR AMPLIFIER FNOOB
;
; DON'T BLAME ME IF AUTO VOLUME UP ON TRACK CHANGE HAPPEN WHEN NOTHING 
; THERE CHANGED -- SEE IT UNPREDICTABLE
;
; THE CODE HERE WILL ALWAYS START PLAYING OF VOLUME UP WHEN RUN
; THE VARIABLE THAT COMPARE HAS TO FILL WITH SOMETHING FIRST OR SET 
; THAT EMPTY YOURSELF TO SOMETHING
;
; DON'T PUT A SCROLLING TITLE IN WINAMP SETTING
;
; USING PRODUCT OF THE FULL VERSION OF V 5.552

; THE TITLE WILL SAY PAUSE STATE WHEN HAPPEN SO THAT WILL 
; SIGNAL A CHANGE OF TRACK -- UNLESS CODE THAT IN

; IT IS NOTICED THE MODULE CODE HAS TO RUN TWICE TO ENABLE IT 
; TO SEE WORKING OF A TRACK CHANGE
; ESPECIALLY WHEN CODING ONLY SOMETIMES HAPPEN
;-----------------------------------------------------------------

;-----------------------------------------------------------------
;-----------------------------------------------------------------
; NEXT UP -- 
;----------- TO LOGGER ALL THE TUNES THAT ARE PLAYED TO TEXT FILE WITH TIME
;----------- VISUAL BASIC FOR ME FOR THAT
;----------- JOLLY DIFFICULT FILE WRITTEN APPEND

;----------- JUST NOTICED AND WANT TO KNOW IF IT IS THIS OR PROJECT ONE THAT AS FOR THE REQUEST FOR WAIT UNTIL TRACK CHANGE HAPPEN -- AS ANOTHER PROGRAM ALSO CONTROLS IT FOR ME
; MIGHT BE HELPFUL
;-----------------------------------------------------------------
;-----------------------------------------------------------------

;-----------------------------------------------------------------
;-----------------------------------------------------------------
; NOTE 
;-----------------------------------------------------------------
; A LOT OF CODERS FOR THIS ONE AUTOHOTKEYS HOOKER DONT USE NOTEPAD VERY WELL 
; TO TAB AND SHIFT - INDENT AND OUT-DENT
; THE CODE THROUGH NESTING AND -- IF -- ELSE -- END IF --
;-----------------------------------------------------------------

;-----------------------------------------------------------------
; WANT ---- WHAT WILL BH HERE IN THE CODE -- IDEA ----
;-----------------------------------------------------------------

;--------------------
#SingleInstance force
;--------------------
;--------------------

;---------------------------------------------------------
; SOURCE CODE FIND -- CREDIT BELONG HERE FOR GETTING STARTED
;---------------------------------------------------------
;---------------------------------------------------------
;---------------------------------------------------------
; The Winamp Song Title & Function Bar - Scripts and Functions - AutoHotkey Community
;https://autohotkey.com/board/topic/30840-the-winamp-song-title-function-bar/
;---------------------------------------------------------
;---------------------------------------------------------
; Get title of WinAmp in AutoHotkey - Google Search
; https://www.google.co.uk/search?num=30&rlz=1C1SKPL_enGB417&q=get+title+of+winamp+in+autohotkey&oq=get+title+of+winamp+in+autohotkey&gs_l=serp.3...143925.145829.0.147106.4.3.0.0.0.0.444.444.4-1.1.0....0...1c.1.64.serp..4.0.0.FkTgqQ9AAmg
;---------------------------------------------------------

#Persistent
DetectHiddenWindows, On
SoundBeep , 1500 , 400

Saved_WINAMP_Title=
setTimer WinampTitle_SUB,100
return

WinampTitle_SUB:
IfWinExist ahk_class Winamp v1.x
{
    WinGetTitle Current_WINAMP_Title,ahk_class Winamp v1.x
    if Current_WINAMP_Title<>%Saved_WINAMP_Title%
    {
        TESTER = Winamp [Paused]
        StringGetPos, pos_PAUSED__1, Current_WINAMP_Title, %TESTER%
        TESTER = Winamp [Stopped]
        StringGetPos, pos_STOPPED_2, Current_WINAMP_Title, %TESTER%
        
        TESTER = Winamp [Paused]
        StringGetPos, pos_PAUSED__3, Saved_WINAMP_Title, %TESTER%
        TESTER = Winamp [Stopped]
        StringGetPos, pos_STOPPED_4, Saved_WINAMP_Title, %TESTER%

        TESTER = (FNOOB TECHNO RADIO)
        StringGetPos, FNOOB_TECHNO_, Current_WINAMP_Title, %TESTER%
    
        TESTER = Winamp [Buffer:
        StringGetPos, pos_Buffer__3, Saved_WINAMP_Title, %TESTER%
	
	

        ;--------------------------------------------------
        ; SPECIAL FOR THE REMEMBERED STATE OF STOPPED 
        ; SEMI IGNORE THE STOP OF REMEMBERED STORED FLAG VAR
        ; AND ALLOW TO RAISE THE VOLUME UP 
        ; AS IF THE TRACK NAME HAD CHANGED
        ; ALLOW BETTER TESTING  1 2 3 4 MSUCIAL SCORE
        ;---------------------------------------------------
        ; IGNORE IN THE LATER NEXT IF LINE STATEMENT CONDITION
        ;---------------------------------------------------
        if pos_STOPPED_4 > 0 
            Saved_WINAMP_Title:=Current_WINAMP_Title
        if FNOOB_TECHNO_ = -1
            Saved_WINAMP_Title:=Current_WINAMP_Title
        ; upto here TIME CHECK 10:20p another 20 minute get this far
        ; AND DONE 
        ;---------
        ; #### /\/\ ******************************************
        ; IT IS MY DELIGHT AND IT'S MY PLEASURE TO OFFER YOUR EDUCATION AND TO ENJOY
        ; FNOOB TECHNO FOR WINAMP V5.552
        ; ASS-HOLE CODING -- BUILD UP MY LEARNING CURVE
        ; #### /\/\ ******************************************

if (FNOOB_TECHNO_ > 0 and pos_Buffer__3 = 0)
{
	; if (pos_STOPPED_2 > 0 or pos_PAUSED__1 > 0)
	if (pos_STOPPED_2 > 0)
	{
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------
		; ____ ^!{Insert} __ Control Alt Insert __ The Winamp By Default Media Global Hot Key is Not Set On
		; ____ I Did Have a Perfect Code But Verbose Backup Was Not Made in Notepad++ and File Overwrite By Other Program 
		; ____ That Was USER
		; ____ Alternative is Hot Key Method 
		; ____ I Have Go It in Clipboard Copy Of Most Important Line and Require Another Extra Second Sub Routine
		; ____ I Was Quite Handsome Learn and Extreme Hard to Find
		; ____ To Set the Winamp Control Play By Code SendMessage Direct Call and Detect If Play Pause Stop
		; ____ As Two Different Function Depend IF Stop OR Pause
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------
		; ____ NEXT PROBLEM AFTER DEBUGGING IDIOT
		; ____ WITH THESE METHOD TEST THIS FUNCTION AND THEN
		; ____ WINAMP DOES NOT ALLOW THE USER OF PAUSE STOP ONLY WHEN IT IS PRESSING AND IN FOCUS 
		; ____ SO YOU HAVE TO MOVE AWAY __ TEMP WORK AROUND 
		; ____ RACE AGAINST TIME TO CONTINUE MY NICE CODE WORKING
		; ____ ANOTHER PROBLEM MOVING AWAY FROM FOCUS CAN ONLY BE THE DESKTOP IS IN FOCUS
		; ____ BIT HARD TO WORK WITH THEN BUT FOR MOMENT MAKE A BEEPER WHEN NOT PLAYING
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------
		;Send, ^!{Home} ; HOME HOT KEY MEDIA KEY WAS FUNCTION OF PAUSE BUT NOT SURE IF THAT WAS UNLOCK HER
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------
		Sendinput ^!{Insert}
		SoundBeep , 1500 , 400
	}
	
	if (pos_PAUSED__1 > 0)
	{
		Sendinput ^!{Insert}
		SoundBeep , 1000 , 400
	}
}

		
        if (FNOOB_TECHNO_ > 0 and pos_PAUSED__1 = -1 and pos_PAUSED__3 = -1 and pos_STOPPED_2 = -1)
        {
            ;-------------------------------------------------------------------
            ; RESULT WORKING 
            ; THE EXTRA MODIFIED HAD TO BE DONE
            ; OF COURSE IT NOW ABLE TO USE ONLY WITH TECHNO INTERNET RADIO FNOOB
            ; AND NOT YOUR USUAL AUDIO SOUND MUSICK LOGICK MAGICK
            ;-------------------------------------------------------------------
            ; EXTRA TIME CODE THAT -- BIT
            ; 10:00 PM
            ;-------------------------------------------------------------------
            Saved_WINAMP_Title:=Current_WINAMP_Title
            SoundBeep , 1500 , 400
            GOSUB, RUN_ME_TO_SET_VOLUME_UP
        }
        ;----------------------------------
        ;`n
        ;HARD CHARACTER TO FIND FOR NEWLINE
        ;----------------------------------
        ;always use bracket around DOUBLE if CHECKER

    }
}
return

RUN_ME_TO_SET_VOLUME_UP:
{
    ; IS HIDDEN A POSSIBLE STATE OF WINAMP
    ; MINIMIZED OR ICON TASKBAR
    ;-----------------------------------
    DetectHiddenWindows, On
    ;-----------------------------------
    IfWinExist ahk_class Winamp v1.x
    {
        SendMessage, 1024, 0, 104   ; IsWinampPlaying?
        WinampPlayStatus := ErrorLevel
        If( WinampPlayStatus == 1 )
        {
            ;----
            ; Winamp: set Volume to a given value - Ask for Help - AutoHotkey Community
            ; https://autohotkey.com/board/topic/53975-winamp-set-volume-to-a-given-value/
            ;----
            ;----
            ; SETVOLUME WINAMP AUTOHOTKEYS HOOKER - Google Search
            ; https://www.google.co.uk/search?q=SETVOLUME+WINAMP+AUTOHOTKEYS HOOKER&rlz=1C1SKPL_enGB417&oq=SETVOLUME+WINAMP+AUTOHOTKEYS HOOKER&aqs=chrome..69i57.23407j0j7&sourceid=chrome&ie=UTF-8
            ; ----
            
            ;------------------------------------------------------------
            ; ABOVE AND BELOW
            ; WITH HELP OF AND MODIFIED BY ME TO MAKE WORK WELL
            ; BY TAKE THE EXAMPLE FROM CODE ABOVE GIVEN IN SOURCE MENTION AT BEGININGER
            ;------------------------------------------------------------
            ; BEING AN IDIOT THIS ONE
            ; PERSONAL COMMENT AND OPINION IN SCIENCE TECHNO FACT - STUPID 
            ; TERRIBLE ENOUGH THAT COMPUTER I SO HARD TO MIX WITH ENGLISH
            ;------------------------------------------------------------
            ; CREDIT TO ORIGINATING SOURCE
            ; GREETING PERHAPS
            ; INTELLECTUAL SOURCE PROPERTY 
            ; COPY STAKE A CLAIM RIGHT
            ; MARK A PITCH
            ; CREDIT WHERE DUE FOR WORK PUT IN
            ; GREETINGS YOUR HONOUR
            ; -----------------------------
            
            ; 
            ; SET MASTER VOLUME OF WINAMP -- RATHER THAN STEP UP DOWN
            ;-------------------------------------------------------
            SendMessage, 1024, Mod( 255, 256 ), 122
            SoundBeep , 1500 , 400
        }
        ELSE
        {
            ;UNUSUAL BEEP FOR TRACK CHANGE AND NOT PLAYING ON LOW VOLUME
            
            SoundBeep , 1500 , 400
            ;----------------------------------------------------------
            ; SoundBeep [, Frequency, Duration]
            ; Frequency
            ; The frequency of the sound, which can be an expression. 
            ; It should be a number between 37 and 32767. 
            ; If omitted, the frequency will be 523.
            ; Duration 150
            ;----------------------------------------------------------

            ;----------------------------------------------------------
            ; INDICATE SIGNAL SIGNIFICANT CHANGE TO TRACK HAPPEN EVEN IF 
            ; NOT DONE ANYTHING LIKE PAUSE IF NOT PLAYING
            ; PAUSE WON'T SHOW A CHANGE ANYWAY -- LEAVE COMMENT HERE
            ; HAVE THOUGHT THIS STUPID THROUGH BEFORE
            ; An NON-OPERATION -- MACHINE ASSEMBLY LANGUAGE MNEMONIC FOR 90 NOP MACHINE CODE HEXADECIMAL 90
            ; ASSEMBLY LANGUAGE THE LOW-LEVEL MOTHER OF ALL LANGUAGES HIGH LEVEL
            ; HEX 90 WHEN YOU WANT TO BRIDGE A CONDITION IF-ER 
            ; WHEN YOU ARE REVERSE ENGINEERING -- DISASSEMBLER DISASSEMBLY
            ; THEY WAY IS BECAUSE -- YOU ENABLE TO DO IT AND RELOCATE SHIFTER
            ; .EXE EXTENSION CODE WAS A TYPE OF IMPROVEMENT OVER THAT AS .COM 
            ; WAS LESS ABLE TO DO SOME RELOCATION

            ;
            ;----------------------------------------------------------
        }
    }
}
return



;----------------------------------------------------------
;----------------------------------------------------------
;----------------------------------------------------------
;----------------------------------------------------------
;----------------------------------------------------------

;CHUCK IN THE JUNK MOST RECENT AT TOP
;MSGBOX --%pos_PAUSED__1% `n--%pos_PAUSED__3% `n--%pos_STOPPED_2% `n--%pos_STOPPED_4%

;----------------------------------------------------------
;----------------------------------------------------------

;X = 0
    ;SetKeyDelay, -1, 0
    ;Loop 
    ;{
        ;X += 1
        
        ;WA_VolSet( newvol=255 ) {
        ;SendMessage, 0x400, Mod( newvol, 256 ), 122, ,ahk_class Winamp v1.x
        ;}
        
        ;{
        ;ControlSend, ahk_parent, {up}  ; volume up
        ;ControlSend, ahk_parent, {Numpad8}  ; volume up
        ;Numpad8
        ;ControlSend, ahk_parent, {up}
        ;}
        
        ;if (X > 100)
        ;{
        ;    SoundBeep , 1500 , 400
        ;    break
        ;}
    ;}

    ;WA_VolSet( newvol=255 ) {
    ;SendMessage, 0x400, Mod( newvol, 256 ), 122, ,ahk_class Winamp v1.x