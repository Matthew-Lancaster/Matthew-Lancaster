@CD \
@TITLE TASKKILL EXPLORER ^& RESET

@ECHO.
@ECHO -------------------------------------
@ECHO -------------------------------------
@ECHO TASKKILL __ EXPLORER ^& RESET HER

@ECHO OFF

@REM __ NOT GOOD IDEA USE /T HERE OR WIPE ALL THING CALLED BY EXPLORR MOSTLY THIS CODE TO RESTORE AFTER

TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE"


@ECHO.
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (5,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.

@ECHO.
TASKKILL /F /IM "EXPLORER*"
@ECHO.
@ECHO -------------------------------------
@ECHO -------------------------------------

@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (5,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)

@ECHO.

EXPLORER

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (5,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)@PAUSE

EXIT