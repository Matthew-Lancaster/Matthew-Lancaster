@ECHO OFF


@CD\
@ECHO ------------------------------------------------------
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@TITLE %~f0
@IF NOT "%1%"=="" (

    Rem ECHO %1 %2 %3 %4 %5 %6 %7 %8 %9
    Rem PAUSE

    @ECHO OFF
    @ECHO -----------------------------
    Rem FOR %%A IN (%*) DO (
        Rem ECHO %%A
    Rem )
    Rem @ECHO -----------------------------

    SET VAR_SET_TO_GO=%1 %2 %3 %4 %5 %6 %7 %8 %9
    ECHO !VAR_SET_TO_GO!


    @TITLE %~f0 __ %1 %2 %3 %4 %5 %6 %7 %8 %9 __
    @ECHO -----------------------------
    @ECHO PROCESS PROGRAM EXE TO PARAM
    @ECHO -----------------------------
    )
@ECHO.
@ECHO OFF

ECHO ON


@SET GoodSync_GSync="C:\Program Files\Siber Systems\GoodSync\gsync.exe"

!GoodSync_GSync! !VAR_SET_TO_GO!

@ECHO.

@ECHO OFF

@REM  ---------------------------------------------------------------
ver | find "10.0." > nul
@IF %errorlevel% EQU 0 @SET WIN10__=YES
ver | find "6.1." > nul
@IF %errorlevel% EQU 0 @SET WIN7___=YES
ver | find "5.1." > nul
@IF %errorlevel% EQU 0 @SET WINXP__=YES
@REM  ---------------------------------------------------------------

@REM  ---------------------------------------------------------------
@IF "%WINXP__%"=="YES" @ECHO WINDOWS XP
@IF "%WIN7___%"=="YES" @ECHO WINDOWS 7
@IF "%WIN10__%"=="YES" @ECHO WINDOWS 10
@REM  ---------------------------------------------------------------
)


@REM 40 SECONDS
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@ECHO OFF

SET TIME_SECOND=40

@IF "%WINXP__%"=="YES" (
@ECHO WINDOWS XP
@REM %TIME_SECOND% SECONDS
FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
FOR /L %%# IN (%TIME_SECOND%,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
)

@IF NOT "%WINXP__%"=="YES" (
	TIMEOUT /T %TIME_SECOND%
)
