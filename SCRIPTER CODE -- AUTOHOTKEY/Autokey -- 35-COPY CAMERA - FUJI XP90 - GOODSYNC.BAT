@REM ----------------------------------------------------------------
@REM __ SEMI DUPLICATE WITH BOTH
@REM __ 
@REM __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA - FUJI XP90 - GOODSYNC.BAT
@REM __ C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB_EXE.BAT
@REM __ 
@REM __ Autokey -- 35-COPY CAMERA - FUJI XP90 - GOODSYNC.BAT
@REM __ 
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 19-Oct-2018 15:05:42 Created Time
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #1 1ST CREATE
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #2 PREVIOUS VERSION BEFORE COPIER
@REM __ TO   ---- Wed 24-Oct-2018 02:18:00 Last Editor #3 CURRENTLY
@REM ----------------------------------------------------------------

@REM -- COPY FROM BOTHER BATCH FILE NOW A TWIN -- NOW A TRIPLET

@REM ----------------------------------------------------------------
@REM -- DESCRIPTION
@REM ----------------------------------------------------------------
@REM -- MAKE FUJI CAMERA MULTI MEDIA MEMORY CARD COPY OVER WHEN 
@REM -- PLUG USB LEAD IN
@REM -- WHAT HAPPEN EXPLORER SEE THE SMB DEVICE 
@REM -- HAS AUTO START TO LOAD THE FOLDER FOR EXPLORER
@REM -- THATS IT
@REM -- THEN AUTOHOTKEYS CODE DETECT THAT WINDOW OPEN
@REM -- AND THE REST COMES HERE
@REM -- IT AUTO LOAD GOODSYNC TO RUN JOB AND THEN AFTE
@REM -- RUN MY FINALIZE-ER FOR RENAME THE PHOTO IMAGES THAT WERE PUT ON 
@REM -- IN TO NICE DATE FORMAT AND EXIF INFO IN FOR MAKE OF CAMERA
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 001 SESSION
@REM ----------------------------------------------------------------
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
@REM -- 002 SESSION
@REM ----------------------------------------------------------------
@REM -- NEW CODE COPY FROM ALMOST SAME BUT WANTED IT OWN
@REM -- BUT MAYBE LATER AFTER GET SYNC-ER GOING
@REM -- DIFFICULT WITH PASSING COMMADN LINE AND ADMINISTRATOR GETTING IN THE WAY
@REM -- CONSTIPATED ADMINISTRATOR
@REM ----------------------------------------------------------------
@REM __ FROM  Wed 24-Oct-2018 00:56:01
@REM __ TO    Wed 24-Oct-2018 02:18:00
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

@CD /D "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY"

@SET TITLE_VAR=FUJI_XP90_CAMERA_RUNNNER__AND_GOODSYNC

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

SET FILENAME_ER="%~dp0GOODSYNC_OUTPUT_FUJI_XP90_CAMERA.TXT"

ECHO.
ECHO ON
@ECHO ----------------------------------------------------------------
@ECHO RUN COMMAND 
@ECHO ----------------------------------------------------------------
ECHO.
CMD /A /C "%GoodSync_GSync% sync "D DSC MC _ EVER _ #2 SMB FUJI XP90">%FILENAME_ER%"
@ECHO OFF
ECHO.
ECHO ----------------------------------------------------------------

@REM ------ 1200 UCS-2 (UTF-16 Little Endian Byte Order)
@REM ----------------------------------------------------------------
@REM SEND OUTPUT FILE FROM GOODSYNC RUN TO COMMAND CONSOLE WINDOW
@REM ----------------------------------------------------------------
ECHO ----------------------------------------------------------------
%ConvertCP% 1200 0 /i %FILENAME_ER%
ECHO ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@SET Analysis_Complete=FALSE
@SET Sync_Runner=FALSE
@SET Another_Job=FALSE

@find "-- Analysis Complete" %FILENAME_ER%>NUL
@if %errorlevel%==0 SET Analysis_Complete=TRUE
@find "Sync: Running:" %FILENAME_ER%>NUL
@if %errorlevel%==0 SET Sync_Runner=TRUE
@find "Another job is Syncing in folder" %FILENAME_ER%>NUL
@if %errorlevel%==0 SET Another_Job=TRUE

@REM DEL %FILENAME_ER%

@ECHO OFF

ECHO.
ECHO ----------------------------------------------------------------
ECHO Analysis_Complete %Analysis_Complete%
ECHO Sync_Runner       %Sync_Runner%
ECHO Another_Job       %Another_Job%
ECHO ----------------------------------------------------------------
ECHO.

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" /QUICK_INTRO_DUMMY_RUN

if "%Sync_Runner%"=="TRUE" (
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO SYNC COMPLETE
	ECHO ----------------------------------------------------------------
	ECHO.
	ECHO ----------------------------------------------------------------
	ECHO RUN COMMAND 
	ECHO C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
	ECHO Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk
	ECHO ----------------------------------------------------------------
	ECHO.
	
	START "" "C:\WINDOWS\EXPLORER.EXE " D:\DSC\2017-Now FUJI XP90\FUJI XP90 IMPORT\DCIM\126_FUJI

	START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk"
)



if "%Analysis_Complete%"=="FALSE" (
	if "%Another_Job%"=="FALSE" (
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO .
		ECHO NOT SURE WHY -- xxxxxxxxxx
		ECHO ----------------------------------------------------------------
	)
	if "%Another_Job%"=="TRUE" (
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS ^& SYNC NOT RUN
		ECHO .
		ECHO ANOTHER JOB IS RUNNING -- ********** 
		ECHO ----------------------------------------------------------------
	)
	
	TIMEOUT /T 8
	ECHO ----------------------------------------------------------------
	ECHO RUN COMMAND 
	ECHO C:\SCRIPTER\SCRIPTER CODE -- GITHUB\
	ECHO BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk %TITLE_VAR%
	ECHO ----------------------------------------------------------------
	"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR%
	TIMEOUT /T 7200
	REM ---- TIMEOUT /T 7200 -- 2 HOURING
)

if "%Analysis_Complete%"=="TRUE" (
	if "%Sync_Runner%"=="FALSE" (
		ECHO ----------------------------------------------------------------
		ECHO ANALYSIS RUN BUT NOT SYNC
		ECHO.
		ECHO MOST LIKEY REASON -- NONE CHANGE HAPPEN
		ECHO ----------------------------------------------------------------
		TIMEOUT /T 4
		ECHO ----------------------------------------------------------------
		ECHO RUN COMMAND 
		ECHO C:\SCRIPTER\SCRIPTER CODE -- GITHUB\
		ECHO BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk %TITLE_VAR%
		ECHO ----------------------------------------------------------------
		"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR%
		TIMEOUT /T 300
		REM 300 = 5 MINUTE
	)
)

if "%Sync_Runner%"=="TRUE" (
	ECHO ----------------------------------------------------------------
	ECHO SYNC COMPLETE
	ECHO ----------------------------------------------------------------
	TIMEOUT /T 30
	ECHO ----------------------------------------------------------------
	ECHO RUN COMMAND 
	ECHO C:\SCRIPTER\SCRIPTER CODE -- GITHUB\
	ECHO BAT 45-SCRIPT RUN GITHUB_AHK_MINIMIZE.ahk %TITLE_VAR%
	ECHO ----------------------------------------------------------------
	"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR%
	TIMEOUT /T 300
	REM 300 = 5 MINUTE
)

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


