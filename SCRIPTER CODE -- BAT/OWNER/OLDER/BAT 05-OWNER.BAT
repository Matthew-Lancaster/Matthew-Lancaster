@setlocal EnableDelayedExpansion
@CD\
@ECHO OFF

@TITLE "%~dpf0"

@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO -- BEGIN TIME -- %DATE%--%TIME%
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------

REM START "%1\* - WORK 1 OF 2 - OWNERSHIP__________" /D"%1" /MIN /NORMAL "ICACLS %1\* /setowner "Everyone" ^>NUL"
REM START "%1\* - WORK 2 OF 2 - GRANT PERMISSION __" /D"%1" /MIN /NORMAL "ICACLS %1\* /t  /grant Everyone:F ^>NUL"

@TITLE "%~dpf0" __ 01 OF 02 __ Icacls %1 /setowner "Everyone" /T /C /L /Q 
Icacls %1 /setowner "Everyone" /T /C /L /Q >NUL
@TITLE "%~dpf0" __ 02 OF 02 __ Icacls %1 /t /Q /grant Everyone
Icacls %1 /t /Q /grant Everyone:F >NUL

@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO -- DONE -- %DATE%--%TIME%
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.
@EXIT