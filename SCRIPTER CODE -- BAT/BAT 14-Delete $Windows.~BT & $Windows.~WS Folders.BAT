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

REM PAUSE
@CD \
REM PAUSE
@ECHO READY __ Safe to delete "\WINDOWS.~BT"
@ECHO.
REM PAUSE

@ECHO.

takeown /F C:\$Windows.~BT\* /R /A /D Y
@ECHO.
@ECHO.
icacls C:\$Windows.~BT\* /T /grant administrators:F 
rmdir /S /Q C:\$Windows.~BT

@ECHO.
@ECHO.

takeown /F C:\$Windows.~WS\* /R /A /D Y
@ECHO.
@ECHO.
icacls C:\$Windows.~WS\* /T /grant administrators:F 
rmdir /S /Q C:\$Windows.~WS


taskkill /F /IM Windows10UpgraderApp.exe /T
takeown /F C:\Windows10Upgrade\* /R /A /D Y
@ECHO.
@ECHO.
icacls C:\Windows10Upgrade\* /T /grant administrators:F 
rmdir /S /Q C:\Windows10Upgrade


echo.
echo DONE THE WORK
echo.

@IF "%1"=="" (
	PAUSE
	REM TIMEOUT /T 40
)

:VERY_ENDER_BLOCK
:ENDE

EXIT
EXIT

REM ----
REM Safe to delete "$WINDOWS.~BT" - Google Search
REM https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=Safe+to+delete+%22%24WINDOWS.%7EBT%22&oq=Safe+to+delete+%22%24WINDOWS.%7EBT%22&gs_l=psy-ab.3..0i30k1.6478.6651.0.6915.2.2.0.0.0.0.78.139.2.2.0....0...1.1.64.psy-ab..0.2.137...0i13i30k1.kjHUIirxFxA
REM --------
REM Can I delete $Windows.~BT & $Windows.~WS folders?
REM http://www.thewindowsclub.com/delete-windows-bt-ws-folders-windows-10
REM ----