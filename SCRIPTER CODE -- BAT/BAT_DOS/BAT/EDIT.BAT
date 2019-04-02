@echo off
rem editmode %1

if "%1"=="C:\CALLERID\CI-EMS.LOG" goto endE
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" goto endE
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" goto endE

rem if "%1"=="C:\CALLERID\2DIARY.TXT" C:\coolcid\2diaryz
if "%1"=="C:\CALLERID\2DIARY.TXT" goto ende2

editmode %1


if "%systemdrive%"=="F:" goto editwin2k
if "%systemdrive%"=="I:" goto editwin2k
if "%systemdrive%"=="H:" goto editwin2k
if "%config%"=="CI2" goto editwin
if "%config%"=="CI" goto editwin
if "%config%"=="DOS" goto editwin
if "%config%"=="WIN" goto editwin
if "%config%"=="ALL" goto editwin
if "%config%"=="CON" goto editwin
if "%config%"=="CON2" goto editwin
if "%config%"=="CON3" goto editwin



editmode %1

if errorlevel 2 dosedit %1
editmode %1
if errorlevel 2 goto endE
if errorlevel 1 vb %1
if errorlevel 1 goto endE
goto endE
:editwin

if errorlevel 2 winedit %1
if errorlevel 2 goto endE
if errorlevel 1 vb %1
if errorlevel 1 goto endE
goto endE

:editwin2k

if errorlevel 2 winedit %1
if errorlevel 2 goto endE
if errorlevel 1 vb %1
if errorlevel 1 goto endE

goto endE

:ENDE

if "%1"=="C:\CALLERID\CIDUD5OD.TXT" C:
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" CD C:\CALLERID
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" CALL PRO CIDUD5OD.TXT
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" C:
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" CD C:\CALLLERID
if "%1"=="C:\CALLERID\CIDUD5OD.TXT" CALL PRO CIDUD5OD.TXT
if "%1"=="C:\CALLERID\CI-EMS.LOG" C:
if "%1"=="C:\CALLERID\CI-EMS.LOG" CD C:\CALLERID
if "%1"=="C:\CALLERID\CI-EMS.LOG" CALL PRO CI-EMS.LOG
if "%1"=="C:\CALLERID\2DOUBLE.TXT" 2sort

:ende2

rem if "%1"=="C:\CALLERID\2DIARY.TXT" C:\HISTORY\BACKDIAR

if "%1"=="C:\CALLERID\2DIARY.TXT" call di.bat

ECHO "%1"     >>C:\CALLERID\1editlst.TXT

