;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 40-AUTO ADD PHOTO NAME MESSENGER FACEBOOK.ahk
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

SETTIMER TIMER_PREVIOUS_INSTANCE,1

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path
SetTitleMatchMode 2  ; ANY PARTIAL

; -------------------------------------------------------------------
; SET THE PATH OF PHOTO FOLDER _ WITHOUT RECURSING SUB-FOLDER IS GOOD IDEA
; TO BE ABLE QUICKLY COUNT THEM VERIFY AND SEE AS ORDER TO GO ONTO FACEBOOK
; AND NOT STRETCH MY CODE TOO MUCH ABOUT WANT RECURSING SUB-FOLDER 
; SINGLE FOLDER ONLY AT THE MOMENT
; -------------------------------------------------------------------
; 01 _ MAKE SIMPLE ADD THE PATH GOING TO USE WITH OR NOT THE END BACKSLASH
FILE_PATH_WILDPATH_JPG=D:\DSC\2015-Now Sony\2019 CyberShot HX60V\DCIM\2019 02 22\
; 02 _ STRIP THE END SLASH OFF IF THERE IS ONE
FILE_PATH_WILDPATH_JPG := regexreplace(FILE_PATH_WILDPATH_JPG, "\\$")
; 03 _ ADD THE \*.JPG
FILE_PATH_WILDPATH_JPG=%FILE_PATH_WILDPATH_JPG%\*.JPG

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
VAR_COUNTER=1

; -------------------------------------------------------------------
; VAR_COUNTER_STOP_AT= NUMBER = STOP __ NOTHING 0 NAUGHT = ALL THE WAY
; -------------------------------------------------------------------
VAR_COUNTER_STOP_AT=200
VAR_COUNTER_STOP_AT+=1 ; ADJUSTMENT MAKE NUMBER EASIER COUNT FROM 0
VAR_COUNTER_STOP_AT=

; NORM WOULD BE TO PUT 1 MAYBE FOR A STARTING
; EMPTY UNLIMITED TO END
VAR_COUNTER_STOP_AFTER=200
VAR_COUNTER_STOP_AFTER=

; EMPTY IGNORE USUALLY 1 EXAMPLE
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
; FACEBOOK_URL_TITLE_1=Facebook - Google Chrome

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

; D:\DSC\2015-Now Sony\2019 CyberShot HX60V\DCIM\2019 02 07 _ DAY AT SEAFRONT WHEN SUNNY AND ROUG


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
		SETTIMER F4,3000
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
		FACEBOOK_TIMER_DELAY=14000
		SETTIMER F4,%FACEBOOK_TIMER_DELAY%
	}
	IF SET_GO=TRUE
	{
		WinActivate ; use the window found above
	}
}

IF SET_GO=FALSE
{
	MSGBOX THE APP DOES NOT FIND THE REQUIREMENT TO RUN `nAS MET BY `n`n%FACEBOOK_URL_TITLE_1% `nWindow Does Not Result to WinExist `n`nAnd Also `n`n%FACEBOOK_URL_TITLE_2% `nDoes Not Result to WinExist `n`nGOING TO EXIT
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

FILE_NAME := % FILE_SCRIPT[VAR_COUNTER]
OF_COUNT=of
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


Sendinput ^a{delete}
Sendinput %VAR_COUNTER% %OF_COUNT% %FILE_SCRIPT_COUNT%`n
Sendinput %OutputVar_1%`n

if POS_VAR>0 
	Sendinput %OutputVar_2%`n

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

IF FACEBOOK_URL_TITLE_2_VAR_SET=TRUE
{
	Sendinput ^{home}
	Sendinput {tab}
	; ---------------------------------------------------------------
	; IMPORTANT TO HAVE BIGGER SLEEP AFTER TAB NEXT ITEM OR IT LAND 
	; ON WRONG LOCATION _ BY TIME NEXT INPUT SENDINPUT HAPPEN
	; ---------------------------------------------------------------
	SLEEP 1500 ; ____ 1 & HALF SECOND DELAY
	Sendinput ^{home}
	PLUG_FLAG=TRUE
}
	
IF FACEBOOK_URL_TITLE_1_VAR_SET=TRUE
{
	Sendinput ^{home}
	Sendinput {tab}{tab}{tab}{tab}{tab}
	SLEEP 1000
}

; -----------------------------------------------------------------------
; RESET TIMER DUE TO ANY DELAY
; LIKE ACTIVATE WINDOW & MORE SLEEP TIMER
; -----------------------------------------------------------------------
SETTIMER F4,off
SLEEP 200
SETTIMER F4,%FACEBOOK_TIMER_DELAY%


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
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process,Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		
		;  ----------------------------------------------------------
		; PROBLEM HERE IF PROGRAM THAT CALL THE BATCH FILE IS KILL SO IS THEN BATCH FILE
		; AND WE GET OVER THAT BY GO EXTRA VIA VBSCRIPT ANOTHER FILE
		; COULD OF RUN A  LOOP AND KILL BUT TRY NOT LOSE OWN ONE FIRST
		; [ Saturday 14:55:00 Pm_02 March 2019 ]
		;  ----------------------------------------------------------

		;  ----------------------------------------------------------
		; OTHER OPTION SET PROCESS KILLER
		;  ----------------------------------------------------------
		; Run, BAT_03_PROCESS_KILLER.BAT /F /IM AutoHotKey.exe /T , , Max
		; Run, %ComSpec% /k ""BAT_03_PROCESS_KILLER.BAT" "/F" "/IM" "AutoHotKey.exe" "/T"" , , Max
		; Process,Close, AutoHotKey.exe
		;  ----------------------------------------------------------
	
		; AUTO GENERATED FILE BY HERE VISUAL BASIC ORIGINAL LONG BEFORE AUTOHOTKEY WANT
		; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe
		; D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
		; -------------------------------------------------------------------
		; AND USED BY HERE
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; -------------------------------------------------------------------
		; [ Saturday 14:52:10 Pm_02 March 2019 ]
		; -------------------------------------------------------------------
		; EDITOR COPY PASTE FROM VBS 39-KILL PROCESS.VBS
		; THIS FILE BECAME USE BY
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; AND THEY USE IT HERE THIS ONE
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return



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
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
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

