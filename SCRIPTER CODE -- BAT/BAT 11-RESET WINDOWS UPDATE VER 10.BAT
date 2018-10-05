@ECHO OFF

REM AFTER INSTALL NEW OPPO SYSTEM OLDER VERSION 1607 RATHER THAN NEWER 1703
REM YOU CAN STILL USE UPDATE TO WINDOWS PROFESSIONAL WHICH TAKES A TIME
REM BUT DOESN'T DO ANY OTHER UPDATER

REM BUT STILL THERE MUST BE TIME TO STOP WINDOWS UPDATE
REM FROM A NEW INSTALL BEFORE ONCE IT HAS BEGUN
REM IF IT HAS BEGUN THERE WILL BE A 4 GIG UPDATE WAITING THAT IS WHEN USE HERE

REM THE SCRIPT CODE COMES FROM TWO LOCATION
REM THE ORIGINAL SOURCE MOT DISPLAYED HERE
REM EMPTIES ALL THE UPDATE
REM AND THEN RESET THE UPDATE SERVICE TO BE GOING AGAIN
REM TO HELP CORRUPTION OF UPDATE

REM BUT WE DON'T WANT IT UPDATE GOING AFTER
REM THERE IS A FAULT WITH VISUAL BASIC CAN'T RUN ANYMORE
REM IT KNOCKS OUT THE SCHEDULER 
REM AND LEFT WITH HARD WORK TO REINSTALL THE OPPO SYSTEM AND TRY NUMEROUS MEASURE
REM TALK ABOUT LEAVING A LAST LEGACY FOR CHILDREN TECHNOLOGY IS SO CUTTING EDGE
REM CAN'T EVEN RUN VBS VB SCRIPTS SOON
REM PASSING ON TO GENERATIONS DOESN'T HAPPEN WITH MICROSOFT

REM ---------------------------------------------------------------------------------
REM SEEM TO WORK REMOVES PENDING UPDATE
REM AND THE LOG I TESTED CERTAIN BIT THING IF INFO WOULD BE LEFT BEHIND BUT NOT EMPTY
REM AND REMOVES UPDATE ON RESTART AND SHUTDOWN BUTTON START MENU CLICKER

@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ ALL BATCH FILE COMMAND SETS 
@REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
@REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
@REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
@REM __ ADD CODE IF REQUIRE
@REM ----------------------------------------------------------------
@REM __ FROM ---- Thu 03-May-2018 01:31:51
@REM __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
@REM ----------------------------------------------------------------


ECHO ____ RESET WINDOWS UPDATE VER 10.BAT ____ 
TITLE %~f0

IF '%1%'=='/ADMIN' (
	@REM :: Check for ADMIN Privileges Again
	>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	@IF '%errorlevel%' GTR '2' (
	ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	ECHO END
	PAUSE
	EXIT /B 1
	)
@ECHO Got ADMIN Privileges __ Requested and Obtained
)
@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	@EXIT /B
)
@ECHO Got ADMIN Privileges

@echo off

set b=0

:bits
set /a b=%b%+1
if %b% equ 3 (
   goto end1
) 
REM _ NERO UPDATE DEPENDANT HAS TO BE STOPPED ALSO OR ELSE YES NO QUESTION
net stop NAUpdate
net stop bits
echo Checking the bits service status.
sc query bits | findstr /I /C:"STOPPED" 
if not %errorlevel%==0 ( 
    goto bits 
) 
goto loop2

:end1
REM cls
echo.
echo Cannot reset Windows Update since "Background Intelligent Transfer Service" (bits) service failed to stop.
echo.
TIMEOUT /T 40
REM pause
exit

:loop2
set w=0

:wuauserv
set /a w=%w%+1
if %w% equ 3 (
   goto end2
) 
net stop wuauserv
echo Checking the wuauserv service status.
sc query wuauserv | findstr /I /C:"STOPPED" 
if not %errorlevel%==0 ( 
    goto wuauserv 
) 
goto loop3

:end2
REM cls
echo.
echo Cannot reset Windows Update since "Windows Update" (wuauserv) service failed to stop.
echo.
TIMEOUT /T 40
REM pause
exit

:loop3
set app=0

:appidsvc
set /a app=%app%+1
if %app% equ 3 (
   goto end3
) 
net stop appidsvc
echo Checking the appidsvc service status.
sc query appidsvc | findstr /I /C:"STOPPED" 
if not %errorlevel%==0 ( 
    goto appidsvc 
) 
goto loop4

:end3
REM cls
echo.
echo Cannot reset Windows Update since "Application Identity" (appidsvc) service failed to stop.
echo.
TIMEOUT /T 40
REM pause
exit

:loop4
set c=0

:cryptsvc
set /a c=%c%+1
if %c% equ 3 (
   goto end4
) 
net stop cryptsvc
echo Checking the cryptsvc service status.
sc query cryptsvc | findstr /I /C:"STOPPED" 
if not %errorlevel%==0 ( 
    goto cryptsvc 
) 
goto Reset

:end4
REM cls
echo.
echo Cannot reset Windows Update since "Cryptographic Services" (cryptsvc) service failed to stop.
echo.

TIMEOUT /T 40
REM pause
EXIT

:Reset
Ipconfig /flushdns
del /s /q /f "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat" 
del /s /q /f "%SYSTEMROOT%\Logs\WindowsUpdate\*"


if exist "%SYSTEMROOT%\winsxs\pending.xml.bak" del /s /q /f "%SYSTEMROOT%\winsxs\pending.xml.bak" 
if exist "%SYSTEMROOT%\winsxs\pending.xml" ( 
    takeown /f "%SYSTEMROOT%\winsxs\pending.xml" 
    attrib -r -s -h /s /d "%SYSTEMROOT%\winsxs\pending.xml" 
    ren "%SYSTEMROOT%\winsxs\pending.xml" pending.xml.bak 
) 

if exist "%SYSTEMROOT%\SoftwareDistribution.bak" rmdir /s /q "%SYSTEMROOT%\SoftwareDistribution.bak"
if exist "%SYSTEMROOT%\SoftwareDistribution" ( 
    attrib -r -s -h /s /d "%SYSTEMROOT%\SoftwareDistribution" 
    ren "%SYSTEMROOT%\SoftwareDistribution" SoftwareDistribution.bak 
) 

if exist "%SYSTEMROOT%\system32\Catroot2.bak" rmdir /s /q "%SYSTEMROOT%\system32\Catroot2.bak" 
if exist "%SYSTEMROOT%\system32\Catroot2" ( 
    attrib -r -s -h /s /d "%SYSTEMROOT%\system32\Catroot2" 
    ren "%SYSTEMROOT%\system32\Catroot2" Catroot2.bak 
) 


REM cls
echo.
echo DONE THE WORK
echo.

@IF "%1"=="" (
	PAUSE
	REM TIMEOUT /T 40
)


:VERY_ENDER_BLOCK

EXIT



REM shutdown -s -t 0


REM ----
REM Unable to disable “Install Windows Updates and ShutDown” option in Start Menu
REM http://windowsdiscussions.com/unable-disable-%93install-windows-updates-shutdown%94-option-start-menu-37771.html
REM --------
REM Reset Windows Update in Windows 10 Windows Update & Activation Tutorials
REM https://www.tenforums.com/tutorials/24742-reset-windows-update-windows-10-a.html
REM ----

