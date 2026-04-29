@CD\
@ECHO ------------------------------------------------------
@TITLE %~f0
@IF NOT "%1%"=="" (

	REM ECHO %1 %2 %3 %4 %5 %6 %7 %8 %9 
	REM PAUSE

	@ECHO OFF
	@ECHO -----------------------------
	FOR %%A IN (%*) DO (
		ECHO %%A
	)
	@ECHO -----------------------------

	@TITLE %~f0 __ %1 __
	@ECHO PROCESS PROGRAM EXE TO KILLER 
	@ECHO -----------------------------
	@ECHO %1
	@ECHO -----------------------------
	@ECHO FORMAT WILL BE EXPECTING
	@ECHO -----------------------------
	@ECHO %1*
	@ECHO -----------------------------
	@ECHO SOMETHING LIKE LIKE LIKE
	@ECHO -----------------------------
	@ECHO %1.EXE
	@ECHO -----------------------------
	)
@ECHO.
@ECHO OFF

REM PAUSE

SET VAR_SET_TO_GO=%1 %2 %3 %4 %5 
REM SET VAR_SET_TO_GO=CALC*

REM PAUSE

REM ---- VAR SET TO GO NOT HERE AND THEN ASKER
REM ------------------------------------------

IF %1.==. (
	@ECHO ENTER THE KILL REQUIREMENT ADD A WILDCARD FOR YOUR OWN
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	TASKKILL /F /IM %INPUT% /T
)

REM ---- VAR SET TO GO HAS A VALUE AND USER
REM ---- TAKEN WILD-CARD OUT TO USE THE VB CODE DO THAT THING 
REM ---- INSTEAD WHEN PASS IT OVER ---- COMMAND LINE
REM ---------------------------------------
IF NOT %1.==. (
	@REM ------------------------------------------------------------------
	@REM THE SOME @ECHO ON HERE IS IGNORE WITHIN THE CONDITION BLOCK
	@REM AND THEN ADDITIONAL ECHO BY COMMAND ECHO TO SHOW EXECUTION COMMAND
	@REM REDUNDANT COMMAND LEFT IN
	@REM OUR OWN VB CODE WILL ADD /F /IM _*_ AND /T IF SELECTED
	@REM WOULD LIKE FOR ABOVE ALO AS THIS A COMMAND INTENDING TO USE
	@REM OR VB SCRIPTOR
	@REM ------------------------------------------------------------------
	@ECHO RUN HERE
	@ECHO ON
	@ECHO TASKKILL %VAR_SET_TO_GO%
	@ECHO.
	TASKKILL %VAR_SET_TO_GO%
	@ECHO OFF
)

@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (4,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.

@REM ------------------------------------------
@REM ____ IF NOT ANY EXPLORER ARE RUNNING AFTER
@TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE"  | find "INFO: No tasks are running" > nul
@IF %errorlevel% EQU 0 GOTO YES
@GOTO NOT
:YES
@ECHO.
TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE" 
@ECHO.
cmd /c EXPLORER
@ECHO.
TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE" 
@ECHO.
@ECHO ---- EXPLORER.EXE RESET AND RELOAD HAPPEN SUCCESSFULLY
:NOT
@ECHO OFF
@REM ------------------------------------------

@ECHO.

@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (40,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.
	
REM PAUSE
REM EXIT

