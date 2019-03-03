@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@CD\
@ECHO ON

@TITLE "%~dpf0"

REM PAUSE
REM __ IF PROBLEM MAYBE FOLDER IS SOME KIND OF LOCK TRY MOVE DLL TO ANOTHER FOLDER MAYBE TESTER
REM -------------------------------------------------------------------------------------------

REGSVR32 "%~dp0\VB6IDEMouseWheelAddin.dll"

REM __ CORRECT DOUBLE ALSO ___ RUN AS ADMIN
REM __ ONCE REGISTERED WILL SHOW WITHER ADDIN

CD /D "%~dp0"
REM REGSVR32 "VB6IDEMouseWheelAddin.dll"


@REM ------------------------------------------------------------------------------------------------------------------------

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

@ECHO OFF
@ECHO.

@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)

@ECHO.

GOTO JUMP_OVER_TEXT
--------------------------------------------------------------------------------------------
OPEN THE __.VPB 
LOOK AT TOP MENU FOR ADD IN
SEE IT THERE
SET TO LOAD ON START UP

FIND OUT SCROLL_BAR ALSO WILL WORK WHEN HOOVER ON THEM MOUSE WHEEL
THE REGISTER THE DLL IS ALSO REQUIRED
D:\VB6\VB-NT\#__ MOUSE_WHEEL_FIX\VB6IDEMouseWheelAddin __ RUN ME.BAT
IF BATCH RUN OKAY IT WILL WAIT AT END RESULT
IF NOT AND QUICK POP GONE THEN ERROR ACCCESS THE DLL LOCK_ED OR FOLDER

RELOCATE THE FOLDER
AND MAYBE EROR MOUSE ADDIN NOT FOUND WANT REMOVE
THAT IS OKAY
MOUSE WHEEL ADD_IN STILL THE
AND HAS TO BE SET AGAIN

[ Wednesday 21:44:00 Pm_10 May 2017 ] The UK
[ Wednesday 20:44:00 Pm_10 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET
[ Thursday 02:15:00 Am_11 May 2017 ] The UK
[ Thursday 01:15:00 Am_11 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET
--------------------------------------------------------------------------------------------
JUMP_OVER_TEXT:
EXIT

