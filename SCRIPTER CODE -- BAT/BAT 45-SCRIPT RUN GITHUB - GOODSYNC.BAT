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


@REM ----------------------------------------------------------------
@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
@REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
@REM ----------------------------------------------------------------

@ECHO OFF

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
IF NOT '%ADMIN_REQUEST%'=='/ADMIN' @ECHO Got ADMIN Privileges -- By Natual

@ECHO OFF

START CMD /K "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"

EXIT 