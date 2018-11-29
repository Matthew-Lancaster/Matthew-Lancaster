@REM ----------------------------------------------------------------
@REM __ C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-SCRIPT RUN GOOD OLE GOODSYNC SCRIPTOR FOLDER.BAT
@REM __ 
@REM __ BAT 59-SCRIPT RUN GOOD OLE GOODSYNC SCRIPTOR FOLDER.BAT
@REM __ 
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 19-Oct-2018 15:05:42 Created Time _ Original Creation for GITHUB
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #1   1ST CREATE 
@REM __ 
@REM __ FROM ---- Wed 07-Nov-2018 20:30:19 Last Editor #002 CHANGE USE HERE
@REM __ TO   ---- Wed 07-Nov-2018 23:34:00 Last Editor #002
@REM ----------------------------------------------------------------

@REM -- COPY FROM BOTHER BATCH FILE NOW A TWIN

@REM ----------------------------------------------------------------
@REM -- 001 SESSION
@REM -- CODER ADD -- MAKE GOODSYNC RUN THE JOB BY COMMAND-LINE 
@REM -- MY FIRST COMMAND-LINE FOR GOODSYNC WORKING WELL
@REM -- GOODSYNC OUTPUT BY COMMANDLINE WAS IN A UNICODE CODEPAGE HAD
@REM -- TO BE DISCOVERED AND CONVERT FORM UNICODE INTO ASCII
@REM -- NOTEPAD++ HELP IN END WHAT TYPE CONVERSION DETECTED IN
@REM -- USUAL CODE THE RUN GOODSYNC AND THEN CHECK THE OUTPUT FOR 
@REM -- THE PROCESS RAN
@REM ----------------------------------------------------------------
@REM __ FROM  Fri 19-Oct-2018 15:05:42
@REM __ TO    Fri 19-Oct-2018 17:30:00
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 002 SESSION
@REM -- LINK TO HERE _ MAKE CODE BECAUSE WAS GOING TO SHUTDOWN RESTART REBOOT ON ANOTHER COMPUTER
@REM -- AND WANTED A QUICKER SYNC-ER OF SCRIPT FOLDER BEFORE SHUT DOWN WITHOUT GOING TO GOODSYNC
@REM -- CODE ADDER
@REM -- LOT OF CHANGES AS USE __ SETLOCAL EnableDelayedExpansion
@REM -- MEAN A LOT OF VARIABLE FORMAT HAS TO CHANGE
@REM -- AND ALSO THE FIND OF STRING WASN'T WORKING TO LOOK UP IN UNICODE OUT 
@REM -- THAT VERY EASY USE NEW COMMAND TO CHANGE TO ANSI FOR SEARCH HERE REQUIRE
@REM -- NOW MY OTHER CODE HAS TO BE UPDATED SAME AS HERE
@REM -- 
@REM -- THE BEGINNER OF VERY GOOD CODE
@REM -- ADD CODE _ EXTRA ARRAY PAIR WITH END OF OPERATION INFO STATUS
@REM -- NOT TAKE LONG WORKED FIRST TIME
@REM ----------------------------------------------------------------
@REM __ FROM  Wed 07-Nov-2018 20:30:19
@REM __ TO    Wed 07-Nov-2018 22:54:00
@REM __ TO    Wed 07-Nov-2018 23:34:00
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM -- Located in Sync Location 1.. 2..
@REM ----------------------------------------------------------------
@REM -- 1..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\
@REM ----------------------------------------------------------------
@REM -- 2..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- GITHUB\
@REM ----------------------------------------------------------------
@REM --- AND 
@REM -- 3.. HERE
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-SCRIPT RUN GOOD OLE GOODSYNC SCRIPTOR FOLDER.BAT
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
@REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
@REM ----------------------------------------------------------------
@ECHO OFF

IF '%1%'=='/ADMIN' (
	@REM :: Check for ADMIN Privileges Again
	>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	@IF '%errorlevel%' GTR '2' SET NOT_GOT_ADMIN=TRUE
)

@IF "%NOT_GOT_ADMIN%"=="TRUE" (
	ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	ECHO END
	PAUSE
	GOTO ENDE
	EXIT /B 1
)

@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)

@ECHO Got ADMIN Privileges

@ECHO OFF
ECHO ----------------------------------------------------------------

@CD /D "C:\SCRIPTER"

@SET TITLE_VAR=GOOD_OLE_GOODSYNC_RUNNER_SCRIPT_FOLDER

@TITLE %TITLE_VAR%

@SET GoodSync_GSync="C:\Program Files\Siber Systems\GoodSync\gsync.exe"

@REM ----
@REM CONVERTCP download | SourceForge.net
@REM https://sourceforge.net/projects/convertcp/
@REM ----
@SET ConvertCP="C:\SCRIPTER\SCRIPTER -- APP TOOL SET\ConvertCP\convertcp.exe"

@REM WELL IT TOOK A WHILE TO FIND CODE PAGE OF TEXT DOCUMENT 
@REM ANSWER CAME FROM A PAG THAT GIVE IT NOTEPAD++ IN BOTTOM RIGHT CORNER TELL THE ANSWER
@REM AND FOUND FROM THERE
@REM ----
@REM text files - Notepad++ can recognize encoding? - Stack Overflow
@REM https://stackoverflow.com/questions/14246956/notepad-can-recognize-encoding/14247070#14247070
@REM ----

SET FILENAME_ER="%~dp0GOODSYNC_SCRIPT_FOLDER_OUTPUT.TXT"
SET FILENAME_ER_2="%~dp0GOODSYNC_SCRIPT_FOLDER_OUTPUT.TXT.TMP"
SET FILENAME_ER_4="GOODSYNC_SCRIPT_FOLDER_OUTPUT.TXT"



@echo off
set sources[1]="C SCRIPTOR __ x _ 7G TO 1X"
set sources[2]="C SCRIPTOR __ x _ 7G TO 2E"
set sources[3]="C SCRIPTOR __ x _ 7G TO 3L"
set sources[4]="C SCRIPTOR __ x _ 7G TO 4G_ GSD"
set sources[5]="C SCRIPTOR __ x _ 7G TO 5P"
set sources[6]="C SCRIPTOR __ x _ 7G TO 8M_ GSD"

set sources[7]="C C DRIVE ROOT CROSS 00 C SCRIPTER FOLDER"
set sources[8]="C C DRIVE ROOT CROSS 00 SCRIPT RUN"
set sources[9]="C C DRIVE ROOT CROSS 1X"
set sources[10]="C C DRIVE ROOT CROSS 2E"
set sources[11]="C C DRIVE ROOT CROSS 3L"
set sources[12]="C C DRIVE ROOT CROSS 4G"
set sources[13]="C C DRIVE ROOT CROSS 5P"
set sources[14]="C C DRIVE ROOT CROSS 8M _ GSD"

set sources[15]="C ISO __ 7G TO 1X"
set sources[16]="C ISO __ 7G TO 2E"
set sources[17]="C ISO __ 7G TO 3L"
set sources[18]="C ISO __ 7G TO 4G"
set sources[19]="C ISO __ 7G TO 5P"
set sources[20]="C ISO __ 7G TO 8M"

set sources[00]=""
set sources[00]=""
set sources[00]=""
set sources[00]=""
set sources[00]=""
set sources[00]=""
set sources[00]=""


set "x=0"

SETLOCAL EnableDelayedExpansion

for /L %%a in (1,1,20) do (

	ECHO !sources[%%a]!
	
	ECHO.
	ECHO ON
	@ECHO ----------------------------------------------------------------
	@ECHO RUN COMMAND 
	@ECHO ----------------------------------------------------------------
	ECHO.
	@REM DEL !FILENAME_ER! >NUL 2>NUL
	@REM CMD /A /C "%GoodSync_GSync% sync !sources[%%a]!>!FILENAME_ER!"
	@ECHO %GoodSync_GSync% sync !sources[%%a]!^>!FILENAME_ER!
	ECHO.
	%GoodSync_GSync% sync !sources[%%a]!>!FILENAME_ER!
	
	@REM ------------------------------------------------------------
	@REM EXAMPLE WITHOUT VARIABLE
	@REM ------------------------------------------------------------
	@REM CMD /A /C "%GoodSync_GSync% sync "C SCRIPTOR __ x _ 7G TO 1X"      >!FILENAME_ER!"

	@ECHO OFF
	ECHO.
	ECHO ----------------------------------------------------------------

	@REM ------ 1200 UCS-2 (UTF-16 Little Endian Byte Order)
	@REM ----------------------------------------------------------------
	@REM SEND OUTPUT FILE FROM GOODSYNC RUN TO COMMAND CONSOLE WINDOW
	@REM ----------------------------------------------------------------
	ECHO ----------------------------------------------------------------
	ECHO ON
	%ConvertCP% 1200 0 /i !FILENAME_ER! /o !FILENAME_ER_2!
	DEL !FILENAME_ER! >NUL 2>NUL
	RENAME !FILENAME_ER_2! !FILENAME_ER_4!
	TYPE !FILENAME_ER!
	
	ECHO ----------------------------------------------------------------
	@REM ----------------------------------------------------------------

	@SET Analysis_Complete=FALSE
	@SET Sync_Runner=FALSE
	@SET Another_Job=FALSE
	@SET Unresolved_Conflicts=FALSE
	@SET Computer_Is_Offline=FALSE
	@SET Change_Happen=TRUE

	@find "-- Analysis Complete" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Analysis_Complete=TRUE
	@REM Sync: Running: ____ DOESN'T APPEAR TO SHOW ALWAYS OR ANYMORE _ CHANGE TO NEXT ONE __ Sync Complete.
	@find "Sync: Running:" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner=TRUE
	@REM == Sync Complete.
	@find "== Sync Complete." !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner=TRUE
	@find "Another job is Syncing in folder" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Another_Job=TRUE
	@find "unresolved conflits" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Unresolved_Conflicts=TRUE
	@find "Analyze has finished with Conflicts" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Unresolved_Conflicts=TRUE
	@find "Computer is offline or turned off" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Computer_Is_Offline=TRUE
	@find "Changes: 0," !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Change_Happen=FALSE
	
	DEL !FILENAME_ER! >NUL 2>NUL

	@ECHO OFF

	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO !sources[%%a]!
	ECHO.
	ECHO Analysis_Complete   !Analysis_Complete!
	ECHO Change_Happen       !Change_Happen!
	ECHO Sync_Complete       !Sync_Runner!
	ECHO Another_Job         !Another_Job!
	ECHO Conflicts           !Unresolved_Conflicts!
	ECHO Computer_Is_Offline !Computer_Is_Offline!
	ECHO ----------------------------------------------------------------
	ECHO.

	REM "C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" /QUICK_INTRO_DUMMY_RUN

	IF "!Computer_Is_Offline!"=="FALSE" (
		if "!Analysis_Complete!"=="FALSE" (
			if "!Another_Job!"=="FALSE" (
				ECHO ----------------------------------------------------------------
				ECHO ANALYSIS ^& SYNC NOT RUN
				ECHO.
				ECHO NOT SURE WHY -- xxxxxxxxxx
				ECHO -------
				---------------------------------------------------------
				set /a "x+=1"
				set STATUS_1[!X!]="ANALYSIS ^& SYNC NOT RUN -- NOT SURE WHY"
				set STATUS_2[!X!]=!sources[%%a]!

			)
			if "!Another_Job!"=="TRUE" (
				ECHO ----------------------------------------------------------------
				ECHO ANALYSIS ^& SYNC NOT RUN
				ECHO.
				ECHO ANOTHER JOB IS RUNNING -- ********** 
				ECHO ----------------------------------------------------------------
				set /a "x+=1"
				set STATUS_1[!X!]="ANALYSIS ^& SYNC NOT RUN -- ANOTHER JOB IS RUNNING"
				set STATUS_2[!X!]=!sources[%%a]!
			)
		)
	)
	
	IF "!Computer_Is_Offline!"=="TRUE" (
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO.
		ECHO Computer_Is_Offline -- xxxxxxxxxx
		ECHO.
		ECHO !sources[%%a]!
		ECHO ----------------------------------------------------------------
		set /a "x+=1"
		set STATUS_1[!X!]="Computer_Is_Offline"
		set STATUS_2[!X!]=!sources[%%a]!
		REM TIMEOUT /T 10
	)

	if "!Analysis_Complete!"=="TRUE" (
		if "!Sync_Runner!"=="FALSE" (
			ECHO ----------------------------------------------------------------
			ECHO ANALYSIS RUN BUT NOT SYNC
			if "!Unresolved_Conflicts!"=="FALSE" (
				ECHO POSSIBLE REASON __ NONE CHANGE HAPPEN
			)
			if "!Unresolved_Conflicts!"=="TRUE" (
				ECHO POSSIBLE REASON __ NONE CHANGE HAPPEN -- BUT -- CONFLICTS ARE FOUND
			)
			ECHO.
			ECHO ----------------------------------------------------------------
			if "!Unresolved_Conflicts!"=="FALSE" (
				REM TIMEOUT /T 4
			)
		)
	)

	if "!Sync_Runner!"=="TRUE" (
		if "!Unresolved_Conflicts!"=="FALSE" (
			ECHO ----------------------------------------------------------------
			ECHO SYNC COMPLETE
			ECHO ----------------------------------------------------------------
			REM TIMEOUT /T 4
			REM ECHO RUN COMMAND 
			REM ECHO C:\SCRIPTER\SCRIPTER CODE -- GITHUB\
			REM ECHO BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk %TITLE_VAR%
			REM ECHO ----------------------------------------------------------------
			REM "C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR%
			REM TIMEOUT /T 300
			REM 300 = 5 MINUTE
		)
	)
	
	if "!Unresolved_Conflicts!"=="TRUE" (
		ECHO ----------------------------------------------------------------
		ECHO Unresolved Conflicts At
		ECHO !sources[%%a]!
		ECHO ----------------------------------------------------------------
		set /a "x+=1"
		set STATUS_1[!X!]="Unresolved Conflicts At"
		set STATUS_2[!X!]=!sources[%%a]!
		TIMEOUT /T 40
	)
	
ECHO.	
ECHO ----------------------------------------------------------------
ECHO ----------------------------------------------------------------
ECHO.	

)

@REM ----------------------------------------------------------------
@REM FALSE FALL FOOLE - STUNT MAN - FALL GUY - SCRIPTING GUY - HEY GUYS
@REM WEEK-DAY - WEAK-DAY
@REM MONTH-DAY
@REM ----------------------------------------------------------------

if !X! gtr 0 @(
	ECHO ----------------------------------------------------------------
	ECHO ATTENTION REQUIRED 
	ECHO ----------------------------------------------------------------
	ECHO.
	for /L %%a in (1,1,!X!) do (
		ECHO !STATUS_1[%%a]!
		ECHO !STATUS_2[%%a]!
		ECHO.
	)
	ECHO ----------------------------------------------------------------
	TIMEOUT /T 80
)


TIMEOUT /T 20


ECHO.
ECHO ----------------------------------------------------------------

:ENDE
EXIT

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM -- NOT WORKING ODD SOD'S -- UNABLE TO CONVERT THE CODEPAGE 
@REM -- USING DOS COMMANDING
@REM ----------------------------------------------------------------
@REM chcp 1200 >nul & cmd /a /c type "%~dp0GOODSYNC_ER_OUTPUT.TXT" >"%~dp0GOODSYNC_ER_OUTPUT_2.TXT"
@REM cmd /a /c type "%~dp0GOODSYNC_ER_OUTPUT.TXT" >"%~dp0GOODSYNC_ER_OUTPUT_2.TXT"
@REM cmd /a /c TYPE "%~dp0GOODSYNC_ER_OUTPUT_2.TXT"

@REM @echo off
@REM >%~dp0GOODSYNC_ER_OUTPUT.TXT (for /F "tokens=1* delims=:" %%G in ('findstr "^" "Sync: Running:"') do echo %%G = "%%H")


REM USING ARRAYS IN BATCH DOS COMMAND PROMPT LEVEL
REM ----
REM How to loop through array in batch? - Stack Overflow
REM https://stackoverflow.com/questions/18462169/how-to-loop-through-array-in-batch
REM ----

