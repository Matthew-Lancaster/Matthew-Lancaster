@COPY "C:\22_LOGON.BAT" "C:\SYSTEM LOGON.TXT" >NUL
@ECHO ----------------------------------------------------------------
@ECHO LOGON EVENTER __ C:\22_LOGON.BAT __COPY TO__ C:\SYSTEM LOGON.TXT
@ECHO ----------------------------------------------------------------

@ECHO OFF

REM "C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_REBOOT_SYSTEM_START_BATCH.BAT"
REM "C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_REBOOT_LOGON_BATCH.BAT"
REM
REM __ How to get the IP Address for your Local Area Connection on Windows Server? - Server Fault
REM __  http://serverfault.com/questions/48567/how-to-get-the-ip-address-for-your-local-area-connection-on-windows-server
SET varcounter=0
SET wmicmd="wmic NICCONFIG WHERE IPEnabled=true GET IPAddress"
FOR /f "tokens=1 delims={, skip=1" %%a in ('%wmicmd%') do call :SETVAR %%a

@SET VAR_DATE=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%--%TIME:~0,8%
REM __ ----------------------------------------------------------------
SET WORD_REPLACE_1=" "
SET WORD_REPLACE_2=0
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
REM __ Command Quote Mark Stripper " "
CALL SET VAR_DATE=%%VAR_DATE:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%

SET varcounter=0
SET wmicmd="wmic NICCONFIG WHERE IPEnabled=true GET IPAddress"
REM FOR /f "tokens=1 delims={, skip=1" %%a in ('%wmicmd%') do call :SETVAR %%a

REM __ ----------------------------------------------------------------
REM __ Delims IS ONLY 1 CHARACTER SEARCH HERE EXAMPLE EXTRACT TEXT BETWEEN 2 Delims FINDER
FOR /f "tokens=2 delims= " %%a IN (%NAME_BODY%) DO SET VAR_DATE=%VAR_DATE% __ %%a __ %IP%

if not exist "C:\C DRIVE ROOT LOGGER"  MKDIR "C:\C DRIVE ROOT LOGGER"

ECHO %VAR_DATE% -- %COMPUTERNAME% -- %IP%
ECHO %VAR_DATE% -- %COMPUTERNAME% -- %IP% >>"C:\C DRIVE ROOT LOGGER\LOGON BOOT TIME IP INFO_%COMPUTERNAME%.TXT"
@ECHO ----------------------------------------------------------------
TITLE LOGON EVENTER __ %VAR_DATE% -- %COMPUTERNAME% -- %IP%


GOTO ENDER

PAUSE
EXIT

:SETVAR
SET /a varcounter=%varcounter% + 1
IF NOT {%1}=={} (
	REM ECHO NIC %varcounter% address is {%1}
	SET NIC%varcounter%=%1
	SET IP=%1 	
		
)
GOTO :eof


:ENDER

@ECHO.

@ECHO OFF
@ECHO.
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (40,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.
	
EXIT
