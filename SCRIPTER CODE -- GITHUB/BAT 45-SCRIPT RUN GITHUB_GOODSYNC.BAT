@REM ----------------------------------------------------------------
@REM __ C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_EXE.BAT
@REM __ 
@REM __ BAT 45-SCRIPT RUN GITHUB_EXE.BAT
@REM __ 
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 19-Oct-2018 15:05:42 Created Time
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #1 1ST CREATE
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #2 CURRENTLY
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
@REM -- Located in Sync Location 1.. 2..
@REM ----------------------------------------------------------------
@REM -- 1..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB_EXE.BAT
@REM ----------------------------------------------------------------
@REM -- 2..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_EXE.BAT
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

@ECHO ON

@CD /D "C:\SCRIPTOR GITHUB"

@SET TITLE_VAR=GIT_RUNNNER__GOODSYNC

@TITLE %TITLE_VAR%

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

@ECHO.
@ECHO ----------------------------------------------------------------------
@ECHO.
CMD /A /C "%GoodSync_GSync% sync "C SCRIPTOR __ y _ 7G __ GITHUB">"%~dp0GOODSYNC_OUTPUT.TXT""
@ECHO.
@ECHO ----------------------------------------------------------------------

@REM ------ 1200 UCS-2 (UTF-16 Little Endian Byte Order)
@REM ---------------------------------------------------
%ConvertCP% 1200 0 /i "%~dp0GOODSYNC_OUTPUT.TXT"

@SET Analysis_Complete=FALSE
@SET Sync_Runner=FALSE
@SET Another_Job=FALSE

@find "-- Analysis Complete" "%~dp0GOODSYNC_OUTPUT.TXT">NUL
@if %errorlevel%==0 SET Analysis_Complete=TRUE
@find "Sync: Running:" "%~dp0GOODSYNC_OUTPUT.TXT">NUL
@if %errorlevel%==0 SET Sync_Runner=TRUE
@find "Another job is Syncing in folder" "%~dp0GOODSYNC_OUTPUT.TXT">NUL
@if %errorlevel%==0 SET Another_Job=TRUE

@ECHO OFF

ECHO.
ECHO ----------------------------------------------------------------------
ECHO Analysis_Complete %Analysis_Complete%
ECHO Sync_Runner      %Sync_Running%
ECHO Another_Job       %Another_Job%
ECHO ----------------------------------------------------------------------
ECHO.

if "%Analysis_Complete%"=="FALSE" (
	if "%Another_Job%"=="FALSE" (
		ECHO ----------------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO .
		ECHO NOT SURE WHY -- xxxxxxxxxx
		ECHO ----------------------------------------------------------------------
	)
	if "%Another_Job%"=="TRUE" (
		ECHO ----------------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO .
		ECHO ANOTHER JOB IS RUNNER -- ********** 
		ECHO ----------------------------------------------------------------------
	)
	
	TIMEOUT /T 8
	"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk" %TITLE_VAR%
	TIMEOUT /T 7200
	REM ---- 2 HOURING
)

if "%Analysis_Complete%"=="TRUE" (
	if "%Sync_Runner%"=="FALSE" (
		ECHO ----------------------------------------------------------------------
		ECHO ANALYSIS RUN BUT NOT SYNC
		ECHO.
		ECHO MOST LIKEY NONE CHANGES
		ECHO ----------------------------------------------------------------------
		TIMEOUT /T 4
		"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk" %TITLE_VAR%
		TIMEOUT /T 120
	)
)

if "%Sync_Runner%"=="TRUE" (
	ECHO ----------------------------------------------------------------------
	ECHO SYNC COMPLETE
	ECHO ----------------------------------------------------------------------
	TIMEOUT /T 8
	"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk" %TITLE_VAR%
	TIMEOUT /T 120
)

@REM ----------------------------------------------------------------
@REM FALSE FALL FOOLE - STUNT MAN
@REM ----------------------------------------------------------------

@ECHO.

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


