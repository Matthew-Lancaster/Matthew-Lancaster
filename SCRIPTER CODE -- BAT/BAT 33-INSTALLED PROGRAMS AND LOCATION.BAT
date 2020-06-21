@ECHO OFF

@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Wed 16-May-2018 02:10:55
@REM __ TO   ---- Wed 16-May-2018 03:32:00 __ 1 HOUR 22 MINUTE _ ONTO NEXT VBSCRIPT 
@REM __ TO   ---- Wed 16-May-2018 04:38:00 __ 2 HOUR 38 MINUTE
@REM __ TO   ---- Wed 16-May-2018 05:05:00 __ 3 HOUR _ LAST BIT 
@REM ----------------------------------------------------------------

@REM ---------------------------------------------------
@REM RESULT GET ALL INSTALLED PROGRAMS
@REM BUT ERROR IT DON'T FIND WINRAR
@REM SEEM TO BE LACK A FEW COMPARED TO  REVO UNINSTALLER
@REM SOME LOCATION ANEMS ARE NOT PRESENT
@REM I ADD IN VBSCRIPT EXTRA FORMATING
@REM ---------------------------------------------------

@REM ----
@REM How to Get List of Installed Programs in Windows 10 – TheITBros
@REM https://theitbros.com/how-to-get-list-of-installed-programs-in-windows-10/
@REM ----

REM ECHO %TEMP%
REM ECHO %~f0
REM ECHO %~dpnx0
REM ECHO %~dp0

IF EXIST "%~dp0\BAT 33-INSTALLED PROGRAMS AND LOCATION" GOTO NOT_CREATE
MKDIR "%~dp0\BAT 33-INSTALLED PROGRAMS AND LOCATION"
:NOT_CREATE

@SET VAR_DATE=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%--%TIME:~0,8%
REM __ ----------------------------------------------------------------
SET WORD_REPLACE_1=" "
SET WORD_REPLACE_2=0
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
CALL SET VAR_DATE=%%VAR_DATE:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%
SET WORD_REPLACE_1=":"
SET WORD_REPLACE_2=_
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
CALL SET VAR_DATE=%%VAR_DATE:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%
SET WORD_REPLACE_1="-"
SET WORD_REPLACE_2=_
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
CALL SET VAR_DATE=%%VAR_DATE:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%

CALL SET VAR_PATH=%~dp0BAT 33-INSTALLED PROGRAMS AND LOCATION\INSTALLED_PROGRAMS_AND_LOCATION__%VAR_DATE%.TXT
ECHO %VAR_PATH%

ECHO ---------------------------------------------------------------->"%VAR_PATH%1"
ECHO %VAR_PATH%>>"%VAR_PATH%1"
ECHO ---------------------------------------------------------------->>"%VAR_PATH%1"
REM wmic product get name,InstallLocation >"%VAR_PATH%2"
wmic product get Description,InstallLocation >"%VAR_PATH%2"

REM CMD /U /C TYPE "%VAR_PATH%1">"%VAR_PATH%"
CMD /C TYPE "%VAR_PATH%1">"%VAR_PATH%"
CMD /C TYPE "%VAR_PATH%2">>"%VAR_PATH%"
DEL "%VAR_PATH%1"
DEL "%VAR_PATH%2"

"C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 22-INSTALLED PROGRAMS AND LOCATION.VBS" "%VAR_PATH%"

START /MAX "C:\Program Files (x86)\Notepad++\Notepad++.exe" "%VAR_PATH%"

EXIT /B
