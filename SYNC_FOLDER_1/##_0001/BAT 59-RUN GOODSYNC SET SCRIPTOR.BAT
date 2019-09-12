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

@REM -- COPY FROM BROTHER BATCH FILE NOW A TWIN

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
@REM -- 003 SESSION
@REM -- ADD CAMERA BACKUP MODE 
@REM -- OTHER DAY START AND THEN THIS PART TO DO
@REM ----------------------------------------------------------------
@REM __ FROM  Sun 21-Jul-2019 17:54:45
@REM __ TO    Sun 21-Jul-2019 18:10:00 -- 15 MINUTE
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
@REM -- 3.. AND HERE
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT
@REM -- 4.. AND HERE PAIR SYNC-ER BOTH HAPPEN
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB - GOODSYNC.BAT
@REM -- C:\SCRIPTOR GITHUB\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB - GOODSYNC.BAT
@REM ----------------------------------------------------------------
@REM -- ALL SYNC-ER NOW ANY MINUTE _ NOW A UNIVERSAL CODE PROGRAM
@REM ----------------------------------------------------------------
@GOTO JUMP_REMMER
REM -----------------------------------------------------------------
REM THE CODE
REM ON THE HOUR 
REM 1..
REM C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.exe
REM THIS CODE WILL EITHER RUN -- 
REM 2..
REM GITHUB - ONLY 
REM WHICH IS HERE RUN
REM BAT 45-SCRIPT RUN GITHUB.BAT
REM ..
REM OR WILL RUN 
REM 3..
REM GOODSYNC SYNC SET AND THEN GITHUB
REM 4..
REM ON THE HOUR IT WILL BE BOTH
REM ..
REM IF IT RUN GOODSYNC SET JOB IT WILL RUN HERE ITSELF ME 
REM BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT
REM 5..
REM -----------------------------------------------------------------
REM ON THE HOUR 
REM 1..
REM C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.exe
REM IS RUN BY HERE
REM Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
REM SUBROUTINE 
REM GITHUB_MIDNIGHT_AND_MIDDAY_TIMER:
REM THIS WILL NOT ACTOR IF GOODSYNC2GO IS RUN BY THE D DRIVE PROCESS
REM -----------------------------------------------------------------
REM Mon 02-Sep-2019 11:45:40
REM -----------------------------------------------------------------
:JUMP_REMMER


@REM ----------------------------------------------------------------
@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
@REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
@REM ----------------------------------------------------------------

@ECHO OFF

ECHO THE ALL THE COMMAND LINE PARAMETER --  ENV VARS
IF NOT "%*"=="" ECHO %*
IF "%*"==""     ECHO -- NONE COMMANDLINE  PARAMETER --  ENV VARS WERE GIVE ----------
ECHO.

ECHO %* | FIND "/ADMIN" >NUL
IF %errorlevel% EQU 0 SET ADMIN_REQUEST=/ADMIN

IF '%ADMIN_REQUEST%'=='/ADMIN' (
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
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN %*", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)

IF '%ADMIN_REQUEST%'=='/ADMIN'     @ECHO Got ADMIN Privileges -- By Push For Them
IF NOT '%ADMIN_REQUEST%'=='/ADMIN' @ECHO Got ADMIN Privileges -- By Natural


@ECHO OFF

ECHO ----------------------------------------------------------------

@REM SET AT TOP
@REM SETLOCAL EnableDelayedExpansion
@REM SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
SETLOCAL EnableDelayedExpansion

REM ---- FILENAME WITHOUT EXTENSION
@SET TITLE_VAR=%~n0
REM ---- FILENAME WITH   EXTENSION
@SET TITLE_VAR=%~n0%~x0
REM ---- FULL SCRIPT PATH
@SET TITLE_VAR=%~f0

REM ----
REM How can I echo a newline in a batch file? - Stack Overflow
REM https://stackoverflow.com/questions/132799/how-can-i-echo-a-newline-in-a-batch-file
REM ----
REM ---- EXAMPLE ---- ECHO READY!\n!STEADY!\n!GO
REM ----
(set \n=^
%=Do not remove this line=%
)

@TITLE !TITLE_VAR!

ECHO READY STEADY GO
ECHO READY!\n!STEADY!\n!GO

ECHO %* | FIND "GOODSYNC_MODE" >NUL
IF %errorlevel% EQU 0 (
REM ECHO !TITLE_VAR! | FIND "BAT 59-RUN GOODSYNC SET SCRIPTOR" >NUL
REM @if !errorlevel!==0 (
	SET MODE_RUN=GOODSYNC
	REM @CD /D "C:\SCRIPTER"
	GOTO START
)

ECHO %* | FIND "GITHUB_MODE" >NUL
IF %errorlevel% EQU 0 (
	SET MODE_RUN=GITHUB
	REM @CD /D "C:\SCRIPTOR GITHUB"
	GOTO START
)


@REM -- COMMAND LINE DEMO
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT CAMERA_BACKUP_MODE


REM SET MODE_RUN=CAMERA_BACKUP
REM GOTO START


ECHO %* | FIND "CAMERA_BACKUP_MODE" >NUL
IF %errorlevel% EQU 0 (
	SET MODE_RUN=CAMERA_BACKUP
	GOTO START
)

IF "!MODE_RUN!"=="" (

	SET MODE_RUN=GITHUB
	REM @CD /D "C:\SCRIPTER"
	GOTO START
)


:START


ECHO %* | FIND "WITH_GITHUB_MODE" >NUL
IF %errorlevel% EQU 0 SET MODE_RUN_WITH_GITHUB=TRUE

REM SET TITLE_RUN_MODE=BAT 45-SCRIPT RUN GITHUB - GOODSYNC
REM SET MODE_RUN=GITHUB

ECHO RUN MODE -- !MODE_RUN!
REM ECHO !TITLE_RUN_MODE!

@SET PATH_GIT="C:\Program Files\Git\bin"
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

@ECHO OFF






@IF NOT EXIST "%~dp0#NFS-%COMPUTERNAME%\" MKDIR "%~dp0#NFS-%COMPUTERNAME%\"

REM Write %~n0 to get the filename without the extension.
SET FILENAME_ER="%~dp0#NFS-%COMPUTERNAME%\%~n0_GOODSYNC_OUTPUT.TXT"
SET FILENAME_ER_2="%~dp0#NFS-%COMPUTERNAME%\%~n0_GOODSYNC_OUTPUT.TXT.TMP"
SET FILENAME_ER_4="%~n0_GOODSYNC_OUTPUT.TXT"
@REM ------------------------------------------------------------
@REM EXAMPLE WITHOUT VARIABLE
@REM ------------------------------------------------------------
@REM CMD /A /C "%GoodSync_GSync% sync "C SCRIPTOR __ y _ 7G __ GITHUB"      >!FILENAME_ER!"

SET LDA=0

IF !MODE_RUN!==GOODSYNC       GOTO GOODSYNC_SET_ARRRAY
IF !MODE_RUN!==GITHUB         GOTO GITHUB___SET_ARRRAY
IF !MODE_RUN!==CAMERA_BACKUP  GOTO CAMERA___SET_ARRRAY

ECHO MUST BE AN ERROR ESCAPE FROM RUN PRESS CLOSE WINDOW -- CONTROL C
PAUSE
GOTO ENDE

:GOODSYNC_SET_ARRRAY
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 1X"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 2E"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 3L"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 4G_ GSD"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 5P"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ x _ 7G TO 8M_ GSD"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 00 C SCRIPTER FOLDER"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 00 SCRIPT RUN"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 1X"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 2E"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 3L"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 4G"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 5P"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C C DRIVE ROOT CROSS 8M _ GSD"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 1X"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 2E"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 3L"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 4G"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 5P"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C ISO __ 7G TO 8M"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 1X"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 2E"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 3L"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 4G"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 5P"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C GOODSYNC __ 7G TO 8M"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D # MY DOC __ 7G TO 4G_1"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D BT Cloud Sync _ 00 _ Caller ID"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D #0 1 INSTALL 7G 4G_1"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D #0 1 INSTALL 7G 8M_1"
SET /A LDA=LDA+1
SET Sources[!LDA!]="VB G7 VB-NT 8M _3_ LONGER"
SET /A LDA=LDA+1
SET Sources[!LDA!]="VB G7 VB-NT 5P _3_ LONGER GSD"
SET /A LDA=LDA+1
SET Sources[!LDA!]="VB G7 VB-NT 4G _3_ LONGER"
SET /A LDA=LDA+1
SET Sources[!LDA!]=""

REM GOTO NEXT_CONTINUE

:GITHUB___SET_ARRRAY
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 -- ROOT 0 COMP_NAMER"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 -- ROOT 1"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 02 SCRIPTER CODE -- GITHUB"
REM SET /A LDA=LDA+1
REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 ELITESPY"
REM SET /A LDA=LDA+1
REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 VB_KEEP_RUNNER"
REM SET /A LDA=LDA+1
REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_05_COMMON"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 02 BAT 59-RUN GOODSYNC SET SCRIPTOR -- VB WANT QUICK CHANGER"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ 02 BAT 59-RUN GOODSYNC SET SCRIPTOR--VB -- 00_Best_VB_01"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ MY WEB 01 MAIN"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB _ MY WEB 02 WORD SET"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ SYNC_FOLDER_#_1_2"
REM GOTO JUMP
	REM SET /A LDA=LDA+1
	REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ SYNC_FOLDER_001_PAIR_01_BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
	REM SET /A LDA=LDA+1
	REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ SYNC_FOLDER_001_PAIR_02_BAT 45-SCRIPT RUN GITHUB - GOODSYNC.BAT"
REM :JUMP
REM SET /A LDA=LDA+1
REM SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ SYNC_FOLDER_Multiple_Thread Port Scanner 02 CON"
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""

REM -----------------------------------------------------------------
REM -- LEAVE UNTIL END AND THEN THE GITHUB SCRIPT WILL RUN FROM GOODSYNC CALL SCRIPT THE 
REM -- BAT 45-SCRIPT RUN GITHUB.BAT
REM -----------------------------------------------------------------
REM -- MAYBE WE DON'T WANT GITHUB UPDATE WHILE GOODSYNC MODE WORKER
REM -----------------------------------------------------------------
IF !MODE_RUN_WITH_GITHUB!==TRUE GOTO WANT_GITHUB
IF !MODE_RUN!==GITHUB           GOTO WANT_GITHUB
GOTO NEXT_CONTINUE

:WANT_GITHUB
	SET /A LDA=LDA+1
	SET Sources[!LDA!]="C SCRIPTOR __ y _ 7G __ GITHUB"

:NEXT_CONTINUE
GOTO NEXT_CONTINUE

:CAMERA___SET_ARRRAY
	SET /A LDA=LDA+1
	SET Sources[!LDA!]="D DSC 1 __  SONY CYBER SHOT MEDIA CARD BACKUP -- =MC - HX60V1"

	SET GoodSync_GSync="C:\GoodSync\x64\gsync.exe"

:NEXT_CONTINUE



REM -----------------------------------------------------------------
REM -- SORT THE EMPTY ARRAY FOR VOID VALUE
REM -----------------------------------------------------------------
SET LDB=0
for /L %%a in (1,1,!LDA!) do (
	IF NOT !sources[%%a]!=="" (
		SET /A LDB=LDB+1
		SET Sources_02[!LDB!]=!sources[%%a]!
	)
)

SET LDA=%LDB%

REM -- TEST DEBUG
REM SET LDA=1

REM -----------------------------------------------------------------
REM TEST ONLY RUN 1 JOB QUICKER DEBUGGING
REM SET LDA=1
REM -----------------------------------------------------------------

for /L %%a in (1,1,!LDA!) do (
	ECHO %%a -- !sources_02[%%a]!
)



SET "X=0"
SET "FAULT_COUNTER=0"
SET Sync_Runner_TOTAL=TRUE

REM -----------------------------------------------------------------
REM -- DON'T REQUIRE HERE AS LONG AS RUN GITHUB SYNC JOB 
REM -- IT HAS SCRIPT THAT DOES IT ALSO
REM -----------------------------------------------------------------
REM "%~dp0\BAT 59-RUN GOODSYNC SET SCRIPTOR-AHK_MINIMIZE.ahk" !TITLE_VAR!  /QUICK_INTRO_DUMMY_RUN
REM -----------------------------------------------------------------

for /L %%a in (1,1,!LDA!) do (

	@ECHO OFF
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO ----------------------------------------------------------------
	ECHO ----------------------------------------------------------------
	ECHO ----------------------------------------------------------------
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO %%a of !LDA!
	ECHO ----------------------------------------------------------------

	REM STRIP QUOTES THIS BIT BEFORE END :"=
	REM NORMAL VAR BE LIKE !Sources_02[%%a]!
	REM CREDIT FINDER AT END OF CODE 
	
	ECHO !Sources_02[%%a]:"=!
	
	@ECHO ON
	@ECHO.
	@ECHO ----------------------------------------------------------------
	@ECHO RUN COMMAND 
	@ECHO ----------------------------------------------------------------
	@ECHO.
	@REM DEL !FILENAME_ER! >NUL 2>NUL
	@REM CMD /A /C "%GoodSync_GSync% sync !Sources_02[%%a]!>!FILENAME_ER!"
	@ECHO %GoodSync_GSync% sync !Sources_02[%%a]!^>!FILENAME_ER!
	@ECHO.
	%GoodSync_GSync% sync !Sources_02[%%a]!>!FILENAME_ER!
	@ECHO OFF
	
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
	
	@SET Analysis_Complete=FALSE
	@SET Sync_Runner=FALSE
	@SET Another_Job=FALSE
	@SET Unresolved_Conflicts=FALSE
	@SET Computer_Is_Offline=FALSE
	@SET Change_Happen=TRUE
	@SET JOB_NOT_FOUND=FALSE

	@find "-- Analysis Complete" !FILENAME_ER!>NUL
	@if !errorlevel!==0 (
		SET Analysis_Complete=TRUE
	)
	@REM Sync: Running: ____ DOESN'T APPEAR TO SHOW ALWAYS OR ANYMORE _ CHANGE TO NEXT ONE __ Sync Complete.
	@find "Sync: Running:" !FILENAME_ER!>NUL
	@if !errorlevel!==0 (
		SET Sync_Runner=TRUE
		SET Analysis_Complete=TRUE
	)
	@REM == Sync Complete.
	@find "== Sync Complete." !FILENAME_ER!>NUL
	@if !errorlevel!==0 (
		SET Sync_Runner=TRUE
		SET Analysis_Complete=TRUE
	)
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
	
	SET FLAG_HERE=FALSE
	@find " not found" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET FLAG_HERE=TRUE
	@IF !FLAG_HERE!==TRUE (
		@find "Error parsing command line: Job" !FILENAME_ER!>NUL
		@if !errorlevel!==0 SET JOB_NOT_FOUND=TRUE
	)
		
	SET Sync_Runner_ONE=TRUE
	@find "Another job is Syncing in folder" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner_TOTAL=FALSE
	@if !errorlevel!==0   SET Sync_Runner_ONE=FALSE
	@find "unresolved conflits" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner_TOTAL=FALSE
	@if !errorlevel!==0   SET Sync_Runner_ONE=FALSE
	@find "Analyze has finished with Conflicts" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner_TOTAL=FALSE
	@if !errorlevel!==0   SET Sync_Runner_ONE=FALSE
	@find "Computer is offline or turned off" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET Sync_Runner_TOTAL=FALSE
	@if !errorlevel!==0   SET Sync_Runner_ONE=FALSE

	@SET FLAG_HERE=FALSE
	@find " not found" !FILENAME_ER!>NUL
	@if !errorlevel!==0 SET FLAG_HERE=TRUE
	@IF !FLAG_HERE!==TRUE (
		@find "Error parsing command line: Job" !FILENAME_ER!>NUL
		@if !errorlevel!==0 SET Sync_Runner_TOTAL=FALSE
		@if !errorlevel!==0   SET Sync_Runner_ONE=FALSE
	)
	
	@REM 021814 Error parsing command line: Job 'C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 -- ROOT' not found
	@REM Job 'C SCRIPTOR __ y _ 7G __ GITHUB _ 00_Best_VB_01 -- ROOT' not found
	
	IF "!Sync_Runner_TOTAL!"=="FALSE"  SET FAULT_FLAG="TRUE"
	IF "!Sync_Runner_ONE!"=="FALSE"    SET /A "FAULT_COUNTER+=1"
	
	@REM PROBLEM OF CONFLICT OF THIS FILE
	@REM ONLY WAY NOW IS TO EXCLUDE IN GOODSYNC	
	DEL !FILENAME_ER! >NUL 2>NUL

	@ECHO OFF

	ECHO ----------------------------------------------------------------
	ECHO %%a of !LDA!
	ECHO ----------------------------------------------------------------
	ECHO !Sources_02[%%a]:"=!
	ECHO ----------------------------------------------------------------
	ECHO Analysis_Complete   !Analysis_Complete!
	ECHO Sync_Runner         !Sync_Runner!
	ECHO Another_Job         !Another_Job!
	ECHO Conflicts           !Unresolved_Conflicts!
	ECHO Computer_Is_Offline !Computer_Is_Offline!
	ECHO Job Not Found       !JOB_NOT_FOUND!
	ECHO ----------------------------------------------------------------
	ECHO.
	
	SET STATUS_4[%%a]="Analysis_Complete   !Analysis_Complete!"
	SET STATUS_5[%%a]="Sync_Runner         !Sync_Runner!"
	SET STATUS_6[%%a]="Another_Job         !Another_Job!"
	SET STATUS_7[%%a]="Conflicts           !Unresolved_Conflicts!"
	SET STATUS_8[%%a]="Computer_Is_Offline !Computer_Is_Offline!"
	SET STATUS_9[%%a]="Job Not Found       !JOB_NOT_FOUND!"

	IF "!Analysis_Complete!"=="FALSE" (
		IF "!Another_Job!"=="FALSE" (
			IF "!Unresolved_Conflicts!"=="FALSE" (
				IF "!Computer_Is_Offline!"=="FALSE" (
					IF "!JOB_NOT_FOUND!"=="FALSE" (
						SET /A "X+=1"
						ECHO ----------------------------------------------------------------
						ECHO ANALYSIS ^& SYNC NOT RUN
						ECHO.
						ECHO NOT SURE WHY                              ---- X-X-X-X-X-X-X-X-X
						ECHO ----------------------------------------------------------------
						ECHO %%a of !LDA!
						ECHO ----------------------------------------------------------------
						SET STATUS_1[!X!]=%%a of !LDA!
						SET STATUS_2[!X!]="ANALYSIS ^& SYNC NOT RUN -- NOT SURE WHY"
						SET STATUS_3[!X!]=!Sources_02[%%a]!
					)
				)
			)
		)
	)
	
	IF "!JOB_NOT_FOUND!"=="TRUE" (
		SET /A "X+=1"
		ECHO ----------------------------------------------------------------
		ECHO JOB NOT FOUND                             ---- X-X-X-X-X-X-X-X-X
		ECHO.
		ECHO ----------------------------------------------------------------
		ECHO %%a of !LDA!
		ECHO ----------------------------------------------------------------
		SET STATUS_1[!X!]=%%a of !LDA!
		SET STATUS_2[!X!]="JOB NOT FOUND ----"
		SET STATUS_3[!X!]=!Sources_02[%%a]!
	)

	IF "!Another_Job!"=="TRUE" (
		SET /A "X+=1"
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO.
		ECHO ANOTHER JOB IS RUNNING                    ---- X-X-X-X-X-X-X-X-X
		ECHO ----------------------------------------------------------------
		ECHO %%a of !LDA!
		ECHO ----------------------------------------------------------------
		SET STATUS_1[!X!]=%%a of !LDA!
		SET STATUS_2[!X!]="ANALYSIS ^& SYNC NOT RUN -- ANOTHER JOB IS RUNNING"
		SET STATUS_3[!X!]=!Sources_02[%%a]!
	)

	IF "!Computer_Is_Offline!"=="TRUE" (
		SET /A "X+=1"
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO.
		ECHO Computer_Is_Offline                       ---- X-X-X-X-X-X-X-X-X
		ECHO.
		ECHO !Sources_02[%%a]:"=!
		ECHO ----------------------------------------------------------------
		ECHO %%a of !LDA!
		ECHO ----------------------------------------------------------------
		SET STATUS_1[!X!]=%%a of !LDA!
		SET STATUS_2[!X!]="Computer_Is_Offline"
		SET STATUS_3[!X!]=!Sources_02[%%a]!
		REM TIMEOUT /T 10
	)

	IF "!Analysis_Complete!"=="TRUE" (
		IF "!Sync_Runner!"=="FALSE" (
			ECHO ----------------------------------------------------------------
			ECHO ANALYSIS RUN BUT NOT SYNC
			IF "!Unresolved_Conflicts!"=="FALSE" (
				ECHO POSSIBLE BECAUSE __ NONE CHANGE HAPPEN
			)
			IF "!Unresolved_Conflicts!"=="TRUE" (
				ECHO POSSIBLE BECAUSE __ NONE CHANGE HAPPEN -- BUT -- CONFLICTS ARE FOUND
			)
			ECHO ----------------------------------------------------------------
			IF "!Unresolved_Conflicts!"=="FALSE" (
				REM TIMEOUT /T 4
			)
		)
	)

	IF "!Sync_Runner!"=="TRUE" (
		IF "!Unresolved_Conflicts!"=="FALSE" (
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
	
	IF "!Unresolved_Conflicts!"=="TRUE" (
		SET /A "X+=1"
		ECHO ----------------------------------------------------------------
		ECHO Unresolved Conflicts At
		ECHO ----------------------------------------------------------------
		ECHO %%a of !LDA!
		ECHO !Sources_02[%%a]:"=!
		SET STATUS_1[!X!]=%%a of !LDA!
		SET STATUS_2[!X!]="Unresolved Conflicts"
		SET STATUS_3[!X!]=!Sources_02[%%a]!
		ECHO ----------------------------------------------------------------
		REM TIMEOUT /T 5
	)

	SET COMP_1="%%a of !LDA!"
	SET COMP_2="!LDA! of !LDA!"

	IF "!Sync_Runner_TOTAL!"=="TRUE" (
		REM ---- FOR THE DURING ONE
		IF NOT "!COMP_1!"=="!COMP_2!" (
			ECHO ----------------------------------------------------------------
			ECHO NOT ERROR OVERALL YET                     ---- *****************
			ECHO ----------------------------------------------------------------
		)
		REM ---- FOR THE LAST ONE
		IF "!COMP_1!"=="!COMP_2!" (
			ECHO ----------------------------------------------------------------
			ECHO NOT ERROR OVERALL TOTALLY SUCCESSFUL      ---- ***************** 
			ECHO ----------------------------------------------------------------
		)
	)
	
	IF NOT "!COMP_1!"=="!COMP_2!" (
		IF "!Sync_Runner_TOTAL!"=="FALSE" (
			ECHO CONFLICT ERROR FOUND OVERALL _ COUNT = !FAULT_COUNTER!  ---- X-X-X-X-X-X-X-X-X
			REM  CONFLICT ERROR FOUND OVERALL _ COUNT = 1  ---- X-X-X-X-X-X-X-X-X
			ECHO ----------------------------------------------------------------
			for /L %%a in (1,1,!X!) do (
				ECHO !STATUS_1[%%a]:"=!
				ECHO !STATUS_3[%%a]:"=!
				ECHO !STATUS_2[%%a]:"=!
				ECHO ----------------------------------------------------------------
			)
			REM TIMEOUT /T 4
		)
	)
	
	ECHO.
)

@REM ----------------------------------------------------------------
@REM FALSE FALL FOOLE - STUNT MAN - FALL GUY - SCRIPTING GUY - HEY GUYS
@REM WEEK-DAY - WEAK-DAY
@REM MONTH-DAY
@REM ----------------------------------------------------------------

ECHO.
ECHO ----------------------------------------------------------------
for /L %%a in (1,1,!LDA!) do (
	ECHO %%a of !LDA!
	ECHO !Sources_02[%%a]:"=!
	ECHO !STATUS_4[%%a]:"=!
	ECHO !STATUS_5[%%a]:"=!
	ECHO !STATUS_6[%%a]:"=!
	ECHO !STATUS_7[%%a]:"=!
	ECHO !STATUS_8[%%a]:"=!
	ECHO !STATUS_9[%%a]:"=!
	ECHO ----------------------------------------------------------------
)

if !X! EQU 0 (
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO SYNC SUCCESSFUL                           ---- ***************** 
	ECHO ----------------------------------------------------------------
	ECHO.
)

if !X! GTR 0 (
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO NOT SUCCESSFUL ATTENTION REQUIRING        ---- X-X-X-X-X-X-X-X-X
	ECHO ----------------------------------------------------------------
	ECHO.
	for /L %%a in (1,1,!X!) do (
		ECHO ----------------------------------------------------------------
		ECHO !STATUS_1[%%a]:"=!
		ECHO !STATUS_2[%%a]:"=!
		ECHO !STATUS_3[%%a]:"=!
		ECHO.
	)
	ECHO ----------------------------------------------------------------
	ECHO.
)

IF !MODE_RUN!==CAMERA_BACKUP (
IF "!Sync_Runner_TOTAL!"=="FALSE" (
	ECHO ----------------------------------------------------------------
	ECHO END RUNNER                                ---- *****************
	ECHO NOT __ GOOD CAMERA_BACKUP                 ---- *****************
	ECHO ATTENTION REQUIRE                         ---- *****************
	ECHO ----------------------------------------------------------------
	ECHO.
	PAUSE
)
)

IF "!Sync_Runner_TOTAL!"=="TRUE" (
	ECHO ----------------------------------------------------------------
	ECHO END RUNNER                                ---- *****************
	ECHO SYNC COMPLETE SUCCESSFUL                  ---- *****************
	ECHO ----------------------------------------------------------------
	REM TIMEOUT /T 8
	REM ECHO ----------------------------------------------------------------
	ECHO RUN COMMAND 
	ECHO %~dp0
	ECHO BAT 59-RUN GOODSYNC SET SCRIPTOR-AHK_MINIMIZE.ahk !TITLE_VAR!
	ECHO ----------------------------------------------------------------
	"%~dp0\BAT 59-RUN GOODSYNC SET SCRIPTOR-AHK_MINIMIZE.ahk" !TITLE_VAR!
	ECHO ----------------------------------------------------------------
	ECHO.
	TIMEOUT /T 300
	REM 300 = 5 MINUTE
)

IF "!Sync_Runner_TOTAL!"=="FALSE" (
	ECHO ----------------------------------------------------------------
	ECHO END RUNNER                                ---- *****************
	ECHO NOT __ SUCCESSFUL ATTENTION REQUIRING     ---- *****************
	ECHO ----------------------------------------------------------------
	ECHO.
	TIMEOUT /T 300
	REM 300 = 5 MINUTE
)




:ENDER
IF NOT "!MODE_RUN!"=="GITHUB_ONLY"    TIMEOUT /T 90

@REM ----------------------------------------------------------------
@REM --- ECHO DE-STUPID-POO
@REM --- ECHO 
@REM --- ECHO DE-STUPID-POO
@REM --- DE-STUPID-POO
@REM --- DE-STUPID-POO OR POO
@REM ----------------------------------------------------------------
@REM FALSE FALL FOOLE - STUNT MAN - FALL GUY - SCRIPTING GUY - HEY GUYS
@REM WEEK-DAY - WEAK-DAY
@REM MONTH-DAY
@REM ----------------------------------------------------------------

@ECHO.
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


REM -----------------------------------------------------------------
REM CREDIT FIND BUT ONLY HELP TO LOCATE IN OWN CODE
REM ----
REM BATCH STRIP %~0 - Google Search
REM https://www.google.com/search?safe=active&ei=YZt7XJTWAdme1fAPo52CmA0&q=BATCH+STRIP+%25%7E0&oq=BATCH+STRIP+%25%7E0&gs_l=psy-ab.3...7068.8831..9110...0.0..0.65.340.6......0....1..gws-wiz.......0i71j0i8i7i30.X6lNLPYEc7I
REM ----
REM ----
REM How do you strip quotes out of an ECHO'ed string in a Windows batch file? - Stack Overflow
REM https://stackoverflow.com/questions/804646/how-do-you-strip-quotes-out-of-an-echoed-string-in-a-windows-batch-file
REM ----
REM -----------------------------------------------------------------


REM USING ARRAYS IN BATCH DOS COMMAND PROMPT LEVEL
REM ----
REM How to loop through array in batch? - Stack Overflow
REM https://stackoverflow.com/questions/18462169/how-to-loop-through-array-in-batch
REM ----


REM ----
REM Batch Script Variables
REM https://www.tutorialspoint.com/batch_script/batch_script_variables.htm
REM ----
REM Example
REM @echo off 
REM set globalvar = 5
REM SETLOCAL
REM set var = 13145
REM set /A var = %var% + 5
REM echo %var%
REM echo %globalvar%
REM ENDLOCAL
REM Few key things to note about the above program.

REM The ‘globalvar’ is defined with a global scope and is available throughout the entire script.

REM The ‘var‘ variable is defined in a local scope because it is enclosed between a ‘SETLOCAL’ and ‘ENDLOCAL’ block. Hence, this variable will be destroyed as soon the ‘ENDLOCAL’ statement is executed.
