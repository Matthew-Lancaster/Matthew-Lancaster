REM @SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@CD \
@ECHO OFF

REM @TITLE "%~dpf0"

PAUSE

REM PAUSE

@REM -----------------
REM PAUSE
@REM -----------------

C:\WINDOWS\system32\regsvr32.exe %1

PAUSE

@REM ------------------------------------------------------------------------------------------------------------------------

REM @SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
REM @FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
REM @FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
REM 	PING -n 2 127.0.0.1 >NUL:)

@ECHO.

PAUSE

REM -----------------------------------------------------------------
REM RELOCATE THE FOLDER
REM AND MAYBE EROR _.DLL ACCESS ERROR WITH FOLDER 
REM AND HAS TO BE SET AGAIN

REM [ Wednesday 21:44:00 Pm_10 May 2017 ] The UK
REM [ Wednesday 20:44:00 Pm_10 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET
REM [ Thursday 02:15:00 Am_11 May 2017 ] The UK
REM [ Thursday 01:15:00 Am_11 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET
REM -----------------------------------------------------------------
