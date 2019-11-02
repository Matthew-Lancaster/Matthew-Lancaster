;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 40-AUTO ADD PHOTO NAME MESSENGER FACEBOOK.ahk
;# __ 
;# __ Autokey -- 40-AUTO ADD PHOTO NAME MESSENGER FACEBOOK.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;# ------------------------------------------------------------------
; DESCRIPTION
; THE CODE WILL ENTER FROM FILENAME ON YOUR COMPUTER FOLDER
; THE PHOTO YOU HAVING EXTRACT THE FILENAME FROM PATH 
; AND ADD THAT IN TO THE DESCRIPTION WHEN YOU CREATE AND ALBUM
; AND ALSO WORK TO EDIT THE ALBUM IF MISTAKE FOUND AFTER PUBLISH
;
; DIFFERENT CODING TO DO THOSE BIT BUT ALL AUTOMATIC RUN AS ONE 
; AND DETECTED WHICH OF TWO
; YOU GOT SUPPLY STARTING VARIABLE COUNTER VALUE
; YOU GOT TO POSITION IN THE FIRST PHOTO DESCRIPTION BOX ON FACEBOOK
; RUN THE CODE AND ALL IS ENTER IN HOPEFULLY WITHOUT LOADING THEM IN ERROR
; BUT IF ERROR PAUSE THE AUTOHOTKEY AND START AGAIN FROM WHERE LEFT OFF
; WHEN REACH THE END IT WILL GIVE EXTRA BEEP CODE TO SIGNAL WHICH 
; IT DO FOR EACH PHOTO ENTER
; AND ALSO EXIT ITSELF
; IT IS SET ON THE F4 FUNCTION HOTKEY BECAUSE THAT HOW ORIGINALLY BEGAN 
; UNTIL WORK DONE _ TO MAKE AUTO TIMER AND NOW CALLED AS A SUBROUTINE BY GOSUB COMMAND
; PRESSING F4 WILL ALSO INSERT EXTRA RUN OF THE ROUTINE AND MIGHT SPEED UP TOO QUICKLY
; OPTION IF THERE TO TURN TIMER OFF AND RUN AS ONLY HOTKEY F4
; GOOD DAY
; IT USE THE SYSTEM OF FILENAME OF PHOTO YOUR HAVING
; IS SET THE DATE IS IN FRONT AND DESCRIPTION COME AFTER THE FIRST UNDERSCORE
; THAT THE TICKET I'M DOING
; YOU CAN PLAY TO ALL YOU HEART CONTENT
; DOWNLOAD AUTOHOTKEYS TINY PROGRAM FOR THE PC MICROSOFT
; LEARN HOW TO SCRIPT
; -------------------------------------------------------------------
; YOU MUST ALSO FOR SAFE-GUARDIAN
; HAVE OPEN THE PAGE OF FACEBOOK WITH YOUR-NAME ABOVE THE URL IN TITLE-BAR 
; TO DO THE INPUT PHOTO IN ALBUM ON CHROME
; AND EXTRA SAFEGUARD YOU GOT TO HAVE EXPLORER OPEN AT THE FILES 
; YOUR PLAN TO USER _ MAKE SURE FULL PATH IS DISPLAYED IN EXPLORER TOP BIT
; BUT YOU CAN EDITOR EASIER REMOVE THAT IF WANT TO 
; JUST SAVES SPRAYING SOME OTHER PROGRAM WITH A LOAD OF HOTKEY INPUT
; -------------------------------------------------------------------
; OVER
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; Spend Time Coding Previous Dated Earlier to Find Not Documented Anywhere Yet
; -------------------------------------------------------------------
; FROM   Tue 09-Oct-2018 21:30:00 __ 
; TO     Wed 10-Oct-2018 07:40:00 __ 10 hours, 10 minutes and 0 seconds
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; Add ExitApp When Conditions Are Not Met To Run with Explain MSGBOX on
; -------------------------------------------------------------------
; FROM   Wed 10-Oct-2018 16:13:43 __ 
; TO     Wed 10-Oct-2018 16:32:00 __ 
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
; OnLine Location 
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 40-Auto Add Photo Name Messenger Facebook.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2040-Auto%20Add%20Photo%20Name%20Messenger%20Facebook.ahk
; ----
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; 001 SESSION
; -------------------------------------------------------------------
; BEEN PROGRAM FOR A BIT
; PRETTY SKILLFUL
; ALL AUTOMATED 
; WORKS ON FUNCTION KEY F4 
; BUT TIMER CALLS GOSUB ON THAT HOTKEY ROUTINE
; ABLE SET SPEED IT ALONG BY PRESS F4 
; OR TURN TIMER OFF AND DO ALL BY KEY
; TODAY SESSION _ WORKING _ WHILE DO BIG UPLOAD PHOTO TO FACEBOOK 
; OF AROUND 500 WHILE FINE TUNE CODE FOR NEXT ROUND
; -------------------------------------------------------------------
; FROM   Tue 09-Oct-2018 21:30:00
; TO     Wed 10-Oct-2018 07:40:00 __ 10 hours, 10 minutes and 0 seconds
;# ------------------------------------------------------------------

;
; -------------------------------------------------------------------
#SingleInstance force
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01-INCLUDE MENU 01 of 03.ahk


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path
SetTitleMatchMode 2  ; ANY PARTIAL

Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav


; -------------------------------------------------------------------
; SET THE PATH OF PHOTO FOLDER _ WITHOUT RECURSING SUB-FOLDER IS GOOD IDEA
; TO BE ABLE QUICKLY COUNT THEM VERIFY AND SEE AS ORDER TO GO ONTO FACEBOOK
; AND NOT STRETCH MY CODE TOO MUCH ABOUT WANT RECURSING SUB-FOLDER 
; SINGLE FOLDER ONLY AT THE MOMENT
; -------------------------------------------------------------------
; MAIN FREQUENT FEW SETTING SOME OTHER LATER IN CODE DOWNA BIT
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; 01 _ MAKE SIMPLE ADD THE PATH GOING TO USE WITH OR NOT THE END BACKSLASH
; -------------------------------------------------------------------
FILE_PATH_WILDPATH_JPG=D:\DSC\2015+SONY\2019 CyberShot HX60V\DCIM\2019 10 24
; -------------------------------------------------------------------
; 02 _ STRIP THE END SLASH OFF IF THERE IS ONE
; -------------------------------------------------------------------
FILE_PATH_WILDPATH_JPG := regexreplace(FILE_PATH_WILDPATH_JPG, "\\$")
; -------------------------------------------------------------------
; 03 _ ADD THE \*.JPG
; -------------------------------------------------------------------
FILE_PATH_WILDPATH_JPG=%FILE_PATH_WILDPATH_JPG%\*.JPG

; -------------------------------------------------------------------
; ENTER THE COUNTER BEGIN NUMBER FOR FACEBOOK PHOTO DESCRIPTION 
; -------------------------------------------------------------------
; START AT VALUE
; NORM WOULD BE 1 FOR 1ST
; VAR_COUNTER_START_AT_VALUE
; -------------------------------------------------------------------
VAR_COUNTER_START_AT_VALUE=1

; -------------------------------------------------------------------
; 0=NEW ALBUM
; 1=EDITOR
NEW_ALBUM_OR_EDITOR_PAGE=0
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

	

; SET DELAY SPEED 
; LESS IMAGE QUICKER SPEED ALLOW
; IF TROUBLE LARGE COUNT IMAGE INCREASE DELAY
; IF EDITOR MODE AFTER PUBLISH IS MORE HARD WORK AND EXTRA 
; DELAY REQUIRE
; FACEBOOK_TIMER_DELAY_NORMAL=1000 -- QUICK WHEN GOT TO DO LESS IMAGE PICTURE 1 SECOND
; FACEBOOK_TIMER_DELAY_NORMAL=3000 -- FEW IMAGE-ER LIKE A COUPLE HUNDRED 3 SECOND
; FACEBOOK_TIMER_DELAY_IN_EDITOR=14000 -- NORMAL FOR MOST - EDITOR REDUCE SPEED REQUIRE
; -------------------------------------------------------------------
FACEBOOK_TIMER_DELAY_NORMAL=500
FACEBOOK_TIMER_DELAY_IN_EDITOR=14000


RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER=TRUE

IF RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
	FACEBOOK_TIMER_DELAY_NORMAL=1




; -------------------------------------------------------------------
; ENTER THE COUNTER BEGIN NUMBER FOR FACEBOOK PHOTO DESCRIPTION 
; AT THE NUMBER NEXT NEEDER TO BE ENTER
; NUMBER 1 FOR USUAL STARTER
; YOU ABLE TO START ANYWHERE IF THE PROCESS WAS RUNNING AND MESSED UP 
; OVER SOMETHING HALF WAY THROUGH _ DUE TO TIMER BEING TOO QUICK FOR 
; SYSTEM OR SOMETHING _ FACEBOOK YOUR PROCESSOR INTERNET _ 
; PROCESS PRIORITY CHROME
; -------------------------------------------------------------------
; START AT VALUE
; NORM WOULD BE 1 FOR 1ST
; VAR_COUNTER_START_AT_VALUE
; -------------------------------------------------------------------
VAR_COUNTER=%VAR_COUNTER_START_AT_VALUE%

; -------------------------------------------------------------------
; VAR_COUNTER_STOP_AT= NUMBER = STOP __ NOTHING 0 NAUGHT = ALL THE WAY
; NORMAL WOULD BE NOTHING
; -------------------------------------------------------------------
VAR_COUNTER_STOP_AT=200
VAR_COUNTER_STOP_AT+=1 ; ADJUSTMENT MAKE NUMBER EASIER COUNT FROM 0
VAR_COUNTER_STOP_AT=

; NORM WOULD BE TO PUT 1 MAYBE FOR A STARTING
; OR NORMAL IS NOTHING
; EMPTY UNLIMITED TO END
VAR_COUNTER_STOP_AFTER=200
VAR_COUNTER_STOP_AFTER=

; EMPTY IGNORE USUALLY 1 EXAMPLE
; NORMAL IS NOTHING
VAR_COUNTER_STOP_AFTER_COUNT=1
VAR_COUNTER_STOP_AFTER_COUNT=
; USER WHEN __ IF VAR_COUNTER_STOP_AFTER

; READ REM LINE SHORTLY BELOW FOR THIS VARIABLE FOR INFO
; USUAL WHEN START A BATCH SET THIS FALSE
; WHEN LOADING BATCHES OF 200 AT A TIME WHEN BIGGER
HAS_FIRST_BATCH_BEEN_DONE_AND_NEXT_SUBSEQUENT_BATCH_DOARH=FALSE





; -------------------------------------------------------------------
; -------------------------------------------------------------------

; GLOBAL FILE_SCRIPT
; GLOBAL FILE_SCRIPT_COUNT

FILE_SCRIPT_COUNT=0
FILE_SCRIPT := Object()

; FILE_SCRIPT := []



; -------------------------------------------------------------------
; REMEMBER TO TURN RECURSE SUBFOLDER OFF WHEN ANY 
; HIDEN IMAGE OR TOTAL COUNT NOT CORRECT
; Loop, Files, %FILE_PATH_WILDPATH_JPG%, R
; -------------------------------------------------------------------

Loop, Files, %FILE_PATH_WILDPATH_JPG%
{
	FILE_SCRIPT[A_Index] := A_LoopFileName
	FILE_SCRIPT_COUNT := A_Index
}	

IF FILE_SCRIPT_COUNT=0
{
	MSGBOX THERE IS NONE FILE COUNT *.JPG `n@ `n%FILE_PATH_WILDPATH_JPG% `n`nGOING TO EXIT
	EXITAPP
	RETURN
}

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

; -----------------------------------------------------------------------
; USE 0 FOR START SEARCH LOCATION DOES LIKE VB INSTRREV IN REVERSE SEARCH  
; -----------------------------------------------------------------------
; STRIP THE END OF THE FULL PATH SO FILENAME IS GONE JUST THE PATH LEFT
; -----------------------------------------------------------------------
SET_String:=SubStr(FILE_PATH_WILDPATH_JPG,1,InStr(FILE_PATH_WILDPATH_JPG,"\",,0)-1)

; -----------------------------------------------------------------------
; WHEN CREATE ALBUM FIRST TIME PAGE
; -----------------------------------------------------------------------
FACEBOOK_URL_TITLE_1=Matthew Lancaster - Google Chrome
; -----------------------------------------------------------------------
; WHEN IN THE ALBUM EDIT PAGE
; -----------------------------------------------------------------------
FACEBOOK_URL_TITLE_2=Facebook - Google Chrome
; 0=NEW ALBUM
; 1=EDITOR
IF NEW_ALBUM_OR_EDITOR_PAGE=1
	FACEBOOK_URL_TITLE_1=%FACEBOOK_URL_TITLE_2%

IF HAS_FIRST_BATCH_BEEN_DONE_AND_NEXT_SUBSEQUENT_BATCH_DOARH=TRUE
	FACEBOOK_URL_TITLE_1=%FACEBOOK_URL_TITLE_2%


; HAS_FIRST_BATCH_BEEN_DONE_AND_NEXT_SUBSEQUENT_BATCH_DOARH=TRUE
; THE THING WITH ADD PHOTO TO ALBUM OF FACEBOOK
; IT DOES NOT LIKE BIG ALBUM ANYMORE
; IT FAVOR YOU GET A FAN PAGE WHERE ABLE TO UPLOAD MANY AS WANT USING AN APP
; THE APP WHICH THEY BANNED FOR PRIVATE ACCOUNT
; SO A SLIGHTLY COMPLEX TASK
; DRAG AND DROP IS FUNNY 
; BECAUSE SOMETIMES SORT ORDER IS NOT THERE AS THE SYSTEM REVERT TO OLD DOS DAY'S 
; WHERE THE SORT ORDER NOT PRESENT
; IT ONLY BECAUSE EXPLORER DOES THE SORT UNDERNEATH IS NONE SORTING
; AND BETTER IS USE A FILE MANAGE TYPE THING LIKE FILELOCATOR PRO / AGENT RANSACK 
; AS WHEN DISPLAY RESULT AND SORT THEM AND THEN DRAG DROP THEM GET BETTER SORT THERE
; NEXT THING
; WHEN OPEN ALBUM TO CREATE NEW
; LOAD IN ONE PICTURE AND THEN THAT TURN ON THE DRAG AND DROP OPTION AFTER
; MAKE SURE GOT CORRECT PRIVILEGED TO EXPLORER DRAG DROP ONTO CHROME WHATEVER AS WINDOWS 10 FUNNY
; YOU REQUIRE SYSTEM OF EXPLORE IS ALWAYS RUN IN PRIVILEGED MODE WHICH HARD THING TO DO NOW-A-DAYS
; AND REQUIRE REGISTRY CHANGE
; I HAVE DETAIL IN GITHUB SOMEWHERE OR SEARCH FOR IT ONLINE HARD ONE GET THERE IN EVENTUALLY 
; REQUIRE SEARCH IN FORUM 1ST & NOT VERY REVEALING UNTIL FOUND WEB SITE HAS IT ALL
; GOT ONE LOADED AND THEN NOW DRAG DROP NEXT ONLY ABOUT 200 ONCE DONE SAVE
; NEXT LOAD THE ALBUM IN EDITOR AND ADD PHOTO PUT ONE IN AGAIN AND DRAG DROP 
; ANOTHER 200 OR 300 BUT RELIABLE LOSS WHEN MORE
; LAST THING AFTER THE FIRST BATCH DONE OF 200
; THE URL WILL CHANGE TO EDITOR AND EDITOR VERSION HAS DIFFERENCE TIMING AND LESS QUICK
; BUT EDITOR URL DOESN'T REALLY MEAN EDITOR MODE IT IS JUST STILL SAME ADD CREATE ALBUM QUICKER AND BETTER
; SO WE GOT THE MANUALLY CHEAT THE SYSTEM A BIT TO TELL NOW 1ST BATCH HAS BEEN DONE
; MAKE CODER _ USE A VARIABLE FOR THAT
; FACEBOOK_URL_TITLE_2=%FACEBOOK_URL_TITLE_1%



; USE SET_GO HERE IF BOTH WINDOW OPEN PRIORITY TO FIRST ONE
; ---------------------------------------------------------
SET_GO=FALSE

; -------------------------------------------------------------------
; TRIM THIS LINE A BIT BECAUSE EXPLORER TITLE DOES THAT ALSO
; SO SEARCH PARTIAL JUST FOR WHAT THE HECK
; -------------------------------------------------------------------

SET_String = % SubStr(SET_String, 1, 95)

; SEARCH FOR EXPLORER WINDOW WITH FILES IN
IfWinExist, %SET_String%
{
	IfWinExist, %FACEBOOK_URL_TITLE_1%
	{
		SET_GO=TRUE
		FACEBOOK_TIMER_DELAY=%FACEBOOK_TIMER_DELAY_NORMAL%
		SETTIMER F4,%FACEBOOK_TIMER_DELAY%
		; MSGBOX %FACEBOOK_TIMER_DELAY%
	}
	
	IF SET_GO=FALSE
	IfWinExist, %FACEBOOK_URL_TITLE_2%
	{
		SET_GO=TRUE
		; -----------------------------------------------------------
		; WHEN IN THE ALBUMS EDIT PAGE _ TIMER FOR UPDATE HAS 
		; TO BE LESS QUICK OR ERROR TAB TO NEXT ONE
		; GIVING ENOUGH SPEED IS CRITICAL TAB WILL JUMP IF NOT READY 
		; FOR NEXT ONE IN
		; -----------------------------------------------------------
		FACEBOOK_TIMER_DELAY=%FACEBOOK_TIMER_DELAY_IN_EDITOR%
		SETTIMER F4,%FACEBOOK_TIMER_DELAY%
	}
	IF SET_GO=TRUE
	{
		IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
			WinActivate ; use the window found above
	}
}

IfWinNOTExist, %SET_String%
{
	MSGBOX YOU HAVE TO BE IN THE EXPLORER FOLDER WHERE FILE EXIST`nTHE APP DOES NOT FIND THE REQUIREMENT TO RUN `nAS MET BY EXIST `n`n%SET_String% `n`nWINDOW DOES NOT RESULT TO WINEXIST`n`nGOING TO EXIT`n`nA BIT FUSSY TRUE
	EXITAPP
	RETURN
}

IF SET_GO=FALSE
{
	MSGBOX THE APP DOES NOT FIND THE REQUIREMENT TO RUN `nAS MET BY `n`n%FACEBOOK_URL_TITLE_1% `nWINDOW DOES NOT RESULT TO WINEXIST`n`nAnd Also `n`n%FACEBOOK_URL_TITLE_2% `n`nDoes Not Result to WinExist `n`nGOING TO EXIT
	EXITAPP
	RETURN
}
Return


; -------------------------------------------------------------------
; HERE FOR FACEBOOK 
; GOT SOME PHOTO IN FOLDER
; THIS WILL GET FILENAME WITH ANY DESCRIPTION
; AND PUT IT IN FACEBOOK 
; SO YOU DON'T HAVE TO TYPE ANY IN ADD YOUR DESCRIPTION ON COMPUTER 
; FIRST AND DON'T HAVE TO MAKE EXTRA COPY
; WISH SPACE BOOK GRIN BOOK _ AI _ WOULD DO SOMETHING LIKE THAT
; THOUGH GUESS THEY ARE SECRETLY KEEPING ALL YOUR FILENAME INFO WITH EACH PHOTO SO IT DON'T GET LOST OR SOMETHING
; -------------------------------------------------------------------
; REM IN AND OUT WHEN WANT TO USE IT 
; SUPPLY PATH NAME _ UP A BIT
; MIGHT MAKE IT OWN AHK FILE THIS ONE AS USED A LOT AND ONLY WANT 
; CALL UP WHEN USER
; MY FAV HOTKEY IS F4 F5 AND OTHER BUT HARDLY USE MANY OR I GET IT TOO COMPLICATED
; GUESSING WHERE MY FAVOURITE HOTKEY ARE
; USE THEM WHEN YOU WANT
; WRITTEN WELL BEFORE THIS DATE
; JUST USE AGAIN TODAY
; -------------------------------------------------------------------
; [ Tuesday 21:31:30 Pm_09 October 2018 ]
; -------------------------------------------------------------------


F4::
{
; F4 -- HERE IS LOAD FROM TIMER AND SHUT OFF AT ENDER
; -------------------------------------------------------------------


FILE_NAME := % FILE_SCRIPT[VAR_COUNTER]
OF_COUNT=of
IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
IF GetKeyState("Capslock", "T")
{
	StringLower, FILE_NAME, FILE_NAME
	StringUpper, OF_COUNT, OF_COUNT
}
ELSE
{
	StringUpper, FILE_NAME, FILE_NAME
	StringLower, OF_COUNT, OF_COUNT
}
	
POS_VAR:=InStr(FILE_NAME,"_")

if POS_VAR>0 
{
	POS_VAR-=1
	OutputVar_1:=SubStr(FILE_NAME, 1, POS_VAR)
	POS_VAR+=2
	OutputVar_2:=SubStr(FILE_NAME, POS_VAR)
}
else
{
	OutputVar_1=%FILE_NAME%
}

StringReplace, OutputVar_2, OutputVar_2,.JPG,,

SoundBeep , 1000 , 150

PLUG_FLAG=FALSE
; JUST IN CASE MOVE
FACEBOOK_URL_TITLE_1_VAR_SET=FALSE
FACEBOOK_URL_TITLE_2_VAR_SET=FALSE

IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
IfWinExist, %SET_String%
{
	IfWinExist, %FACEBOOK_URL_TITLE_1%
	{
		WinActivate ; use the window found above
		PLUG_FLAG=TRUE
		FACEBOOK_URL_TITLE_1_VAR_SET=TRUE
	}
	IF PLUG_FLAG=FALSE
		IfWinExist, %FACEBOOK_URL_TITLE_2%
		{
			WinActivate ; use the window found above
			FACEBOOK_URL_TITLE_2_VAR_SET=TRUE
		}
}

END_MESSENGER=Hitt Push the Back Link for Whole Album Top Left

IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
IF GetKeyState("Capslock", "T")
{
	 Lab_Invert_Char_Out:= ""
	 Loop % Strlen(END_MESSENGER) {
		Lab_Invert_Char:= Substr(END_MESSENGER, A_Index, 1)
		if Lab_Invert_Char is upper
		   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) + 32)
		else if Lab_Invert_Char is lower
		   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) - 32)
		else
		   Lab_Invert_Char_Out:= Lab_Invert_Char_Out Lab_Invert_Char
	 }
	 END_MESSENGER=%Lab_Invert_Char_Out%
}
; Source Creditor
; ----
; Convert text - uppercase, lowercase, capitalized or inverted - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/
; ----





IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
{
	Sendinput ^a{delete}
	Sendinput %VAR_COUNTER% %OF_COUNT% %FILE_SCRIPT_COUNT%`n
	Sendinput %OutputVar_1%`n
}

IF RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
{
	VAR_CLIPBOARD=%VAR_CLIPBOARD%%VAR_COUNTER% %OF_COUNT% %FILE_SCRIPT_COUNT%`n%OutputVar_1%`n
}


IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
if POS_VAR>0 
	Sendinput %OutputVar_2%`n


	
	
; Sendinput %END_MESSENGER%`n

	
; -------------------------------------------------------------------
; IF ALREADY PUBLISHED PHOTO AND WANT TO ADD INFO DESCRIPTION 
; USE ONE TAB INSTEAD 
; EDIT ALBUM MODE VS CREATE ALBUM MODE
; CREATE ALBUM MODE _ OPERATES A LOT QUICKER AND LESS FUSSY
; THE CODE AND TIMER LENGTH DELAY GOT PROPER WHAT TO DO FOR EACH ONE
; -------------------------------------------------------------------
; %FACEBOOK_URL_TITLE_2% = EDIT ALBUM MODE
; %FACEBOOK_URL_TITLE_1% = CREATE ALBUM MODE
; -------------------------------------------------------------------


IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
IF FACEBOOK_URL_TITLE_2_VAR_SET=TRUE
{
	Sendinput ^{home}
	Sendinput {tab}
	; ---------------------------------------------------------------
	; IMPORTANT TO HAVE BIGGER SLEEP AFTER TAB NEXT ITEM OR IT LAND 
	; ON WRONG LOCATION _ BY TIME NEXT INPUT SENDINPUT HAPPEN
	; ---------------------------------------------------------------
	; SLEEP 1500 ; ____ 1 & HALF SECOND DELAY
	SLEEP 500 ; ____ 1 & HALF SECOND DELAY
	
	Sendinput ^{home}
	PLUG_FLAG=TRUE
}
	

IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
IF FACEBOOK_URL_TITLE_1_VAR_SET=TRUE
{
	Sendinput ^{home}
	Sendinput {tab}{tab}{tab}{tab}{tab}
	; SLEEP 1000
	SLEEP 500
}


IF !(A_ComputerName = "4-ASUS-GL522VW") 
	Soundplay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav

; -----------------------------------------------------------------------
; RESET TIMER DUE TO ANY DELAY
; LIKE ACTIVATE WINDOW & MORE SLEEP TIMER
; -----------------------------------------------------------------------
IF !RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
{
	SETTIMER F4,off
	; SLEEP 200
	SLEEP 100
	SETTIMER F4,%FACEBOOK_TIMER_DELAY%
}

VAR_COUNTER+=1
VAR_COUNTER_STOP_AFTER_COUNT+=1

STOP_HERE=FALSE
IF VAR_COUNTER>%FILE_SCRIPT_COUNT%
	STOP_HERE=TRUE
	
IF VAR_COUNTER_STOP_AT
	IF VAR_COUNTER>%VAR_COUNTER_STOP_AT%
		STOP_HERE=TRUE

IF VAR_COUNTER_STOP_AFTER
	IF VAR_COUNTER_STOP_AFTER_COUNT
		IF VAR_COUNTER_STOP_AFTER_COUNT>%VAR_COUNTER_STOP_AFTER%
			STOP_HERE=TRUE
		
IF STOP_HERE=TRUE
{
	IF RESULT_CLIPBOARD_NOT_ENTER_AT_BROWSER
	{
	clipboard = %VAR_CLIPBOARD%
	}

	SETTIMER F4,off
	SoundBeep , 1000 , 150
	SoundBeep , 2000 , 200
	SoundBeep , 1000 , 150
	ExitApp
	; ---------------------------------------------------------------
	; ALL COMPLETE AND THEN EXIT AND MAKE SOUND TO NOTIFY
	; ---------------------------------------------------------------
}
}
RETURN


MenuHandler:
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02-INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03-INCLUDE MENU 03 of 03.ahk



;# ------------------------------------------------------------------
; USUAL END BLOCK OF CODE TO HELP EXIT ROUTINE
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
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

;# ------------------------------------------------------------------
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
;# ------------------------------------------------------------------
; exit the app

