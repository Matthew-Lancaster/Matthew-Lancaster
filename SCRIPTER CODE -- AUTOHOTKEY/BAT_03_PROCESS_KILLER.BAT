@ECHO OFF
REM -- AUTO GENERATED FILE BY HERE VISUAL BASIC
REM -- D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe
REM -- D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
REM -- -------------------------------------------------------------------
REM -- AND USED BY HERE
REM -- LOT OF AUTOHOTKEYS TRAY MENU ITEM
REM -- -------------------------------------------------------------------
REM -- [ Saturday 14:52:10 Pm_02 March 2019 ]
REM -- -------------------------------------------------------------------
REM -- EDITOR COPY PASTE FROM VBS 39-KILL PROCESS.VBS
REM -- THIS FILE BECAME USE BY
REM -- LOT OF AUTOHOTKEYS TRAY MENU ITEM
REM -- AND THEY USE IT HERE THIS ONE
REM -- C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
REM -- ORIGINAL AT HERE LOCATION 
REM -- C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
REM -- AND MOVED HERE MAYBE 
REM -- -------------------------------------------------------------------
REM -- MOST LIKELY TRY AND KEEP IN SYNC LATER
REM -- EXCEPT THE AUTO GENERATOR
REM -- -------------------------------------------------------------------

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
    @ECHO PROCESS PROGRAM EXE TO KILLER
    @ECHO -----------------------------
    )
@ECHO.
@ECHO OFF



Rem ---- VAR SET TO GO NOT HERE AND THEN ASKER
Rem ------------------------------------------

IF %1.==. (

    @ECHO ENTER THE KILL REQUIREMENT ADD A WILDCARD FOR YOUR OWN __ EXAMPLE __ TASKKILL /F /IM INPUTNAME /T
    @SET INPUT=
    @SET /P INPUT=Type input: %=%
    ECHO.
    ECHO --------------------------------------------------
    ECHO RUN COMMAND LINE ____ /F /IM !INPUT! /T
    ECHO.
    TASKKILL /F /IM !INPUT! /T
)

Rem ---- VAR SET TO GO HAS A VALUE AND USER
Rem ---- TAKEN WILD-CARD OUT TO USE THE VB CODE DO THAT THING
Rem ---- INSTEAD WHEN PASS IT OVER ---- COMMAND LINE
Rem ---------------------------------------
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
    @ECHO -----------------------------
    @ECHO ON
    @ECHO TASKKILL %VAR_SET_TO_GO%
    @ECHO.
    TASKKILL %VAR_SET_TO_GO%
    @ECHO OFF
)

@REM __ IF EXPLORER ASK TO BE KILL WAIT FEW SECOND
@REM __ CHECKER AND RESUME

@ECHO OFF

ECHO %1 %2 %3 %4 %5 %6 %7 %8 %9 | findstr /I "EXPLORER.EXE"
if %errorlevel% == 0 SET EXPLORER_METHOD_TO_GO__=YES

IF NOT "%EXPLORER_METHOD_TO_GO__%"=="YES" (
@ECHO.
@ECHO --------------------------------
@ECHO EXPLORER USE KILLER NOT DETECTOR
)

IF "%EXPLORER_METHOD_TO_GO__%"=="YES" (
    @ECHO.
    @ECHO EXPLORER DETECTOR WAIT FEW MOMENT....
    @ECHO.

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
: YES
    @ECHO.
    TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE"
    @ECHO.
    @REM -----------------------------------
    cmd /c EXPLORER
    @REM -----------------------------------
    @ECHO.
    TASKLIST /FI "IMAGENAME EQ EXPLORER.EXE"
    @ECHO.
    @ECHO ---- EXPLORER.EXE RESET AND RELOAD HAPPEN SUCCESSFULLY
    :NOT
    @ECHO OFF
    @REM ------------------------------------------
)
@ECHO.

@REM AS NORMAL

@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
    PING -n 2 127.0.0.1 >NUL:)
@ECHO.

REM PAUSE
Rem EXIT
