@echo off
rem editmode %1

if "%1"=="C:\COOLCID\CI-EMS.LOG" goto endE
if "%1"=="C:\COOLCID\CIDUD5OD.TXT" goto endE

if "%1"=="C:\COOLCID\2DIARY.TXT" C:\coolcid\2diaryz

editmode %1


if "%systemdrive%"=="E:" goto editwin2k
if "%config%"=="CI2" goto editwin
if "%config%"=="CI" goto editwin
if "%config%"=="DOS" goto editwin
if "%config%"=="WIN" goto editwin
if "%config%"=="ALL" goto editwin
if "%config%"=="CON" goto editwin
if "%config%"=="CON2" goto editwin



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

if "%1"=="C:\COOLCID\CIDUD5OD.TXT" C:
if "%1"=="C:\COOLCID\CIDUD5OD.TXT" CD C:\COOLCID
if "%1"=="C:\COOLCID\CIDUD5OD.TXT" CALL PROT CIDUD5OD.TXT
if "%1"=="C:\COOLCID\CI-EMS.LOG" C:
if "%1"=="C:\COOLCID\CI-EMS.LOG" CD C:\COOLCID
if "%1"=="C:\COOLCID\CI-EMS.LOG" CALL PROT CI-EMS.LOG
if "%1"=="C:\COOLCID\2DOUBLE.TXT" 2sort
if "%1"=="C:\COOLCID\2DIARY.TXT" C:\HISTORY\BACKDIAR

ECHO "%1"     >>C:\coolcid\1editlst.TXT

