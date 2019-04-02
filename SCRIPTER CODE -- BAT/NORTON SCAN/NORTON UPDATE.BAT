@ECHO off
CD\
ECHO.
REM __ Sun 15 October 2017 21:23:39----------

REM __ ---------------------------
REM __ FIND IN THE PATH TO EXECUTE
REM __ ---------------------------


REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /A
REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /update
REM "C:\Program Files\Norton Internet Security\Engine\22.7.0.76\Navw32.exe" /update /automation
REM @SETLOCAL EnableDelayedExpansion

REM PROBLEM FINDING TWO
REM OVERWRITE OR EXIT LOOP 
REM ANSWER OVERWRITE 

for /f "delims=" %%a in ('dir "C:\Program Files (x86)\Norton Security with Backup\Navw32.exe" /S /B') do (
	REM  call set NORTON=%%NORTON%%%%a
	call set NORTON=%%a
)

IF NOT "%NORTON%"=="" GOTO NEXT_LINE

ECHO OFF
for /f "delims=" %%a in ('dir "C:\Program Files (x86)\Navw32.exe" /S /B') do (
	REM  call set NORTON=%%NORTON%%%%a
	call set NORTON=%%a
)

:NEXT_LINE

ECHO %NORTON%

ECHO ON
START "" "%NORTON%" /update

REM DIR %NORTON%

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