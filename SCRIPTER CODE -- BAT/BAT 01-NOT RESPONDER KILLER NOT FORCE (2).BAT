@ECHO OFF
ECHO ______________________________
ECHO __ MATTHEW LANCASTER
REM ECHO __ MATT.LAN@BTINTERNET.COM
REM ECHO __ Tue 21 February 2017 10:50:00 -- ALL-NIGHTER PREP THIS LEVEL ALL DEBUG WORKING
REM ECHO __ ONE ROUNTINE FINISH DO I'M KNACKERING
ECHO ______________________________

@ECHO CD /D D:\
@CD /D D:\
@ECHO OFF
@REM ------------------------------------------------
TITLE ____ BOOT_KILLER ____
@REM ------------------------------------------------

@ECHO OFF

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

@ECHO.

@REM setlocal EnableDelayedExtensions ____ THIS LINE HAS ERROR SETTER
setlocal EnableExtensions

@ECHO OFF

REM ECHO -----------------------------------------------------------------------------------------------------------------------------------
REM ECHO -- FOR EVENT NOT RESPONDER NEED CRASHER
REM ECHO -----------------------------------------------------------------------------------------------------------------------------------
REM ------------------------------------------------------------------------------------------------
REM ECHO 1.. TASKKILL /F /FI "STATUS EQ NOT RESPONDING" _______ WILD CARD ALL NOT RESPONDING
ECHO 1.. TASKKILL /FI "STATUS EQ NOT RESPONDING" _______ WILD CARD ALL NOT RESPONDING
REM ECHO -----------------------------------------------------------------------------------------------------------------------------------
REM __ ECHO 2.. TASKKILL /FI "IMAGENAME NE Z" /IM * /T     _______ WILD CARD EVERYTHING __ AND __ NOT FORCE 
REM __ ECHO 3.. TASKKILL /F /FI "IMAGENAME NE Z" /IM * /T  _______ WILD CARD EVERYTHING __ YES FORCE BSD BSD XBOX
REM __ ECHO 4.. TASKKILL /FI "WINDOWTITLE NE Z" /IM * /T      ____ WILD CARD EVERYTHING ALL WINDOWTITLE NOT FORCE
REM __ ECHO 5.. TASKKILL /F /FI "WINDOWTITLE NE Z" /IM * /T ______ WILD CARD EVER ALL WINDOWTITLE YES FORCE BSD
REM ------------------------------------------------------------------------------------------------
@SET VAR_COUNTER=21
@SET ACCUMULATOR =0

REM ECHO.
REM ECHO -----------------------------------------------------------------------------------------------------------------------------------
REM ECHO -----------------------------------------------------------------------------------------------------------------------------------
REM ECHO.


REM CALL :WILD_CARD__01_OF_05
@SET ACCUMULATOR =0
REM ECHO ON
REM ------------------------------------------------------------------------------------------------

CALL :WILD_CARD__01_OF_05
GOTO JUMPER_ROUTINE_WILD_CARD__01_OF_05
::
:WILD_CARD__01_OF_05
@REM setlocal
@ECHO OFF
REM -----------------------------------------------------------------------------------------------
REM ECHO ------------------------------------------------------------------------------------------------
REM ECHO.
REM ECHO -- WILD CARD __ 01 OF 05
REM ------------------------------------------------------------------------------------------------
REM -- NOT RESPONDER -------------------------------------------------------------------
REM -- NOT REQUIRE A DOUBLE QUESTION EXAMPLE SHOWN GIVE
REM -------------------------------------------------------------------------------------------------
SET /A ACCUMULATOR=ACCUMULATOR+1
SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
ECHO -----------------------------------------------------------------------------------------------
ECHO.
ECHO %IX1%
ECHO ON
TASKLIST /FI "STATUS EQ NOT RESPONDING"
@ECHO OFF
TASKLIST /FI "STATUS EQ NOT RESPONDING"  | find "INFO: No tasks are running" > nul
IF %errorlevel% EQU 0 GOTO NOT

ECHO.
ECHO __ Do You Want Kill ---- ALL NOT RESPONDING SERVICE SET AS SHOWN AND OR PROCESS SET 
ECHO __ WILL MAYBE KILLER CMD.COM THIS BAT SCRIPT LEARN 
ECHO __ ITSELF ONLY __ IF SLACKING INTO NOT RESPONDING
ECHO.

ECHO __ 20 SECOND AND NOT RESPONDING PROCESS ABOVE WILL BE KILLER GENITALLY __ EXIT IF WANT TO
ECHO.

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (10,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.

Set INPUT=Y
IF "%INPUT%" == "y" GOTO YES
IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
ECHO.
ECHO ON
REM TASKKILL /F /FI "STATUS EQ NOT RESPONDING"
TASKKILL /FI "STATUS EQ NOT RESPONDING"
@ECHO OFF
:NOT
ECHO.
endlocal & goto :EOF
@REM REMMER --------------------------------------------------------------------------------
:JUMPER_ROUTINE_WILD_CARD__01_OF_05


ECHO -----------------------------------------------------------------------------------------------
ECHO.
REM ------------------------------------------------------------------------------------------------
REM -- FINALLY EDITOR NOTEPAD LEARN 
REM ------------------------------------------------------------------------------------------------
SET /A ACCUMULATOR=ACCUMULATOR+1
ECHO -----------------------------------------------------------------------------------------------
ECHO.
REM ECHO %ACCUMULATOR% OF %VAR_COUNTER%  -- DO YOU WANT EDIT HERE NOTEPAD++ C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT 
REM ECHO ____________________ ?(Y/N)
REM SET INPUT=
REM SET /P INPUT=Type input: %=%
REM IF "%INPUT%" == "y" GOTO YES
REM IF "%INPUT%" == "Y" GOTO YES
REM GOTO NOT
REM :YES
REM START "" /MAX /HIGH "C:\Program Files (x86)\Notepad++\notepad++.exe" "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT"
REM GOTO ENDING
REM :NOT
REM ECHO.

ECHO.
ECHO -----------------------------------------------------------------------------------------------

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.
	
REM EXIT