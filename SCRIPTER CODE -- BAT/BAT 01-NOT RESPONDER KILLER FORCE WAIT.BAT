@ECHO OFF
REM ______________________________
REM __ MATTHEW LANCASTER
REM ______________________________
REM ECHO __ MATT.LAN@BTINTERNET.COM
REM ECHO __ Tue 21 February 2017 10:50:00 -- ALL-NIGHTER PREP THIS LEVEL ALL DEBUG WORKING
REM ECHO __ ONE ROUTINE FINISH DO I'M KNACKERING
REM -----------------------------------------------------------------
REM SESSION 002
REM -----------------------------------------------------------------
REM WORK DO FROM -- Thu 23-Jan-2020 11:33:43
REM WORK DO TO   -- Thu 23-Jan-2020 14:02:00 -- 2 HOUR 28 MINUTE
REM -----------------------------------------------------------------
REM "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER FORCE WAIT.BAT"
REM "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER FORCE.BAT"
REM "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER NOT FORCE WAIT.BAT"
REM "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER NOT FORCE.BAT"
REM -----------------------------------------------------------------

ECHO CD /D D:\
CD /D D:\

REM ---------------------------------------------------------------
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET WIN10__=YES
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET WIN07__=YES
ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET WINXP__=YES
REM ---------------------------------------------------------------

REM ---------------------------------------------------------------
IF "%WINXP_%"=="YES" ECHO WINDOWS XP
IF "%WIN7__%"=="YES" ECHO WINDOWS 7
IF "%WIN10_%"=="YES" ECHO WINDOWS 10
REM ---------------------------------------------------------------

SETLOCAL ENABLEEXTENSIONS

SET DESCRIPTION=TASK KILLER NOT RESPOND FORCE AND WAIT
TITLE ____ %DESCRIPTION% ____
ECHO.
ECHO %DESCRIPTION%

SET CMD_TASKKILLER=/F /FI "STATUS EQ NOT RESPONDING"
SET CMD_TASKLIST__=/FI "STATUS EQ NOT RESPONDING"

ECHO.
ECHO TASKKILL %CMD_TASKKILLER% _______ WILD CARD ALL NOT RESPONDING
ECHO ----------------------------------------------------------------

REM __ ECHO 2.. TASKKILL /FI "IMAGENAME NE Z" /IM * /T     _______ WILD CARD EVERYTHING __ AND __ NOT FORCE 
REM __ ECHO 3.. TASKKILL /F /FI "IMAGENAME NE Z" /IM * /T  _______ WILD CARD EVERYTHING __ YES FORCE BSD BSD XBOX
REM __ ECHO 4.. TASKKILL /FI "WINDOWTITLE NE Z" /IM * /T      ____ WILD CARD EVERYTHING ALL WINDOWTITLE NOT FORCE
REM __ ECHO 5.. TASKKILL /F /FI "WINDOWTITLE NE Z" /IM * /T ______ WILD CARD EVER ALL WINDOWTITLE YES FORCE BSD
REM ------------------------------------------------------------------------------------------------

ECHO -----------------------------------------------------------------------------------------------
ECHO.
ECHO TASKLIST %CMD_TASKLIST__% _______ WILD CARD ALL NOT RESPONDING
ECHO.
SET FILENAME_ER="%~dp0BAT 01-NOT RESPONDER KILLER NOT FORCE WAIT_OUT.TXT"
TASKLIST %CMD_TASKLIST__% >%FILENAME_ER% 2>>&1
TYPE %FILENAME_ER%
find "INFO: No tasks are running" %FILENAME_ER%>NUL
find "Session Name" %FILENAME_ER%>NUL
SET ERR_1=%ERRORLEVEL%
ECHO.
IF "%ERR_1%"=="0"      ECHO ---- ERRORLEVEL FROM FIND GIVER RESULT ^> 0
IF NOT "%ERR_1%"=="0"  ECHO ---- ERRORLEVEL FROM FIND RESULT NOTHING 0
IF "%ERR_1%"=="0"      SET TIMER_DELAY=20
IF NOT "%ERR_1%"=="0"  SET TIMER_DELAY=4
IF NOT "%ERR_1%"=="0"  GOTO NOT

ECHO.
ECHO __ Do You Want Kill ---- ALL NOT RESPONDING SERVICE SET AS SHOWN AND OR PROCESS SET 
ECHO __ WILL MAYBE KILLER CMD.COM THIS BAT SCRIPT LEARN 
ECHO __ ITSELF ONLY __ IF SLACKING INTO NOT RESPONDING
ECHO.
ECHO __ NOT RESPONDING PROCESS ABOVE WILL BE KILLER FORCE __ EXIT IF WANT
ECHO.
PAUSE

SET INPUT=Y
IF "%INPUT%" == "y" GOTO YES
IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
ECHO.
ECHO ON
TASKKILL %CMD_TASKKILLER%
@ECHO OFF
:NOT

ECHO.
ECHO -----------------------------------------------------------------------------------------------

IF "%WINXP_%"=="YES" (
	@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
	@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
	@FOR /L %%# IN (%TIMER_DELAY%,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
		PING -n 2 127.0.0.1 >NUL:)
)
IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
	GOTO NOT_WINDOWS_XP_JUMP
:NOT_WINDOWS_XP
	TIMEOUT /T %TIMER_DELAY%
:NOT_WINDOWS_XP_JUMP
ECHO.
	
REM EXIT