@ECHO off
CD\
ECHO.

ECHO ____ NORTON UPDATE and QUICK SCAN.BAT ____
SET TITLE_VAR=%~f0
TITLE %TITLE_VAR% - RUN BATCH

REM __ EDIT THE FONT OF %COMSPEC% HERE BY DOING THIS
REM __ START "" "%COMSPEC%"

REM __ Sun 15 October 2017 21:23:39----------

REM __ ---------------------------
REM __ FIND IN THE PATH TO EXECUTE
REM __ ---------------------------

REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /A
REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /update
REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /update /automation

@SETLOCAL ENABLEDELAYEDEXPANSION

SET "NORTON="

for /f "delims=" %%a in ('dir "C:\Program Files (x86)\Norton Security with Backup\Navw32.exe" /S /B') do (
 call set NORTON=%%a
)

IF NOT "!NORTON!"=="" GOTO JUMP_HER
for /f "delims=" %%a in ('dir "C:\Program Files\Norton Security with Backup\Navw32.exe" /S /B') do (
 call set NORTON=%%a
)
:JUMP_HER


ECHO __ DOUBLE SAFTEY CHECKER 
ECHO __ IF NORTON CHANGED THE NAME OF PATH PRESUME STILL IN 32 BIT (x86) PROGRAM FILES
IF NOT "!NORTON!"=="" GOTO JUMP_HER
for /f "delims=" %%a in ('dir "C:\Program Files\Navw32.exe" /S /B') do (
 call set NORTON=%%a
)
:JUMP_HER

REM for /f "delims=" %%a in ('dir "C:\Program Files (x86)\Norton Security with Backup\uiStub.exe" /S /B') do (
 REM call set NORTON_2=%%NORTON_2%%%%a
REM )
REM uiStub.exe /LU /S+


ECHO.
IF NOT "!NORTON!"=="" (
	ECHO !NORTON!
)
ECHO.


REM START "" /WAIT "%NORTON_2%" /LU
REM IF EXIST !NORTON! (
IF NOT "!NORTON!"=="" (
	ECHO WAIT FOR UPDATE TO RUN AND THEN AFTER RUN A QUICK SCAN

	REM ECHO WONT TAKE LONG ESPICALLY WITHOUT UPDATE THE RUN HAPPENS IN THE BACKGROUND
	REM ECHO.
	REM ECHO COMBINE  /UPDATE /QUICK IN ONE COMMAND LINE QUICK SCAN WILL HAPPEN WITH GUI UNLESS TOLD NOT TO WITH 
	REM ECHO /SESCAN QUICK SCAN BACKGROUND
	REM ECHO REQUIRE TO PRESS OKAY AFTER UPDATE __ MAYBE WANT IT SEPERATE OR MINIMAL
	ECHO.

	REM START "" /WAIT "!NORTON!" /UPDATE /SESCAN
	rem START "" "!NORTON!" /B+ /SE+ /L /SESCAN

	REM ECHO MINIMIZE	
	REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

	ECHO START "" "!NORTON!" /UPDATE /QUICK
	START "" "!NORTON!" /UPDATE /QUICK
)


IF "!NORTON!"=="" (
	ECHO THE NORTON SEARCH EXE FILE DID NOT EXIST TO RUN
)


IF NOT "!NORTON!"=="" (
	ECHO.
	ECHO NORTON UPDATE HAS BEGAN AND OKAY BUTTON WILL ARRIVE SHORTLY UNLESS UPDATE TO DO
	ECHO.
	ECHO AUTOHOTKEYS WILL TAKE CONTROL OF THE RUNNER _ REST OF IT LEARN
	ECHO.

	REM THE BATCH HERE IS RUN FROM AUTOHOTKEYS SCRIPT BELOW NOW
	REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 18-NORTON CONTOL BOOTER.ahk"
)
REM START "" "!NORTON!" /QUICK


IF NOT "!NORTON!"=="" (
	ECHO.
	ECHO MINIMIZE THIS WINDOW HAPPEN IN 30 SECOND OR OTHERWIZE
	ECHO AFTER SCAN BEGAN AND _ AFTER UPDATE
	ECHO.
	TIMEOUT /T 30
	TIMEOUT /T 20
)
IF "!NORTON!"=="" (
	TIMEOUT /T 20
)

EXIT
EXIT

REM ECHO ON
REM TIMEOUT /T 5
REM @ECHO OFF
REM ECHO.
REM ECHO MINIMIZE THIS WINDOW ONLY AS LONG AS STILL IN FOCUS
REM ECHO.
REM ECHO ON
REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"d"
REM @ECHO OFF
REM ECHO.
REM TIMEOUT /T 20


REM ECHO MINIMIZE	
REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"
REM ECHO MINIMIZE	
REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM START "" /WAIT "!NORTON!" /UPDATE
REM THOUGHT YOU COULD RUN BOTH ONE COMMADN LINE IN TURN
REM BUT HAS ERROR OF SCAN CANCELLED
REM START "" "!NORTON!" /B+ /SE+ /QUICK /SESCAN


REM TIMEOUT /T 10

REM START "" "!NORTON!" /QUICK

REM DIR "!NORTON!"

REM ECHO OFF

REM PAUSE
REM -----
REM -----

REM PROGRAM OPTIONS _ DOUBLE CLICK TO FIND
REM ----
REM How to use NAVW32 utility?
REM https://support.norton.com/sp/en/us/norton-security-with-backup/22.11.0.41/solutions/v10817483_ns_retail_en_us?abproduct=Norton+Security+with+Backup&abversion=22.11.0.41&actstat=activated&buildname=Retail&conntype=1000000000&coreservice=Startup+Type%3Aauto+State%3ARunning&cpu=Intel64+Family+6+Model+94+Stepping+3&datetime=10-15-2017+20%3A38%3A26+GMT&defbrowser=Chrome&dsfree=85.19&dstotal=187.11&eapenabled=false&endpointid=571E548D-ECC0-4A10-BCF9-0DA991E4D345&entsrc=help&env=prod&hbguid=571E548D-ECC0-4A10-BCF9-0DA991E4D345&hcmode=false&heartbeatID=571E548D-ECC0-4A10-BCF9-0DA991E4D345&helpid=IDH_NAVW_COMMAND_LINE_SWITCHES&ieversion=9.11.15063.0&layout=Retail&layouttype=ESD&lic_attr=16928786&lic_type=16&memload=18&memtotal=27371&os=windows&oslang=iso%3AENG&oslocale=iso%3AGBR&osvers=10.0&osversion=10.0+15063.0.amd64fre.rs2_release.170317-1834&partnerid=1000&partnername=&plang=sym%3AEN&plgid=30&plid=765&psn=XKFW6C9YDJP4&puid=5039&remdays=603&skuf=21352088&skum=21363030&skup=21351089&spversion=0.0&sublength=846&subremaining=603&substatus=current&symskucurrent=21351089&symskumedia=21363030&utm_medium=product&utm_source=symc&vendorid=1000
REM ----


REM DIR AND OUTPUT TO VARIABLE - Google Search
REM https://www.google.co.uk/search?q=DIR+AND+OUTPUT+TO+VARIABLE&rlz=1C1CHBF_en-GBGB759GB759&oq=DIR+AND+OUTPUT+TO+VARIABLE&aqs=chrome..69i57j0l5.5371j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM How can I save output of dir in a variable?
REM https://www.computerhope.com/forum/index.php?topic=65683.0
REM ----