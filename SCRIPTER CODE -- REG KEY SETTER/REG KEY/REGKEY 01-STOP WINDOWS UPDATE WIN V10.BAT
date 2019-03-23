@ECHO OFF

REM -- [ Sunday 06:10:30 Am_18 November 2018 ] 

REM ----
REM How to Disable Windows 10 Update in Every Way - EaseUS
REM https://www.easeus.com/todo-backup-resource/how-to-stop-windows-10-from-automatically-update.html
REM ----
REM -- 7. Name the new key AUOptions and press Enter.
REM -- Double-click the newly created key and change its value to 2. It's for "Notify for download and notify for install". Click OK.

REM ----
REM Is NoAutoRebootWithLoggedOnUsers still supported/effective on Windows 10? - Super User
REM https://superuser.com/questions/1066441/is-noautorebootwithloggedonusers-still-supported-effective-on-windows-10
REM ----
REM ----
REM Help Windows 10: How to prevent automatic restart for installing updates [TurboLab.it]
REM https://turbolab.it/manutenzione-156/guida-windows-10-come-impedire-riavvio-automatico-installazione-aggiornamenti-843
REM ----

REM C:\PStart\# NOT INSTALL REQUIRED\Reg Converter\RegConvert.exe

REM -- --------------------------------------------------------------
REM -- WHAT HAPPEN HERE IS TRUE WINDOWS UPDATE ASK YOU BEFORE INSTALL
REM -- BUT BIG MAIN UPDATE STILL HAPPEN WITHOUT ASKING
REM -- --------------------------------------------------------------
REM -- [ Monday 07:09:20 Am_18 February 2019 ]
REM -- --------------------------------------------------------------

ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10

IF NOT %VARLATCH2%==WIN10 GOTO ENDE

@ECHO ON
CD\

Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "NoAutoRebootWithLoggedOnUsers" /t REG_DWORD /d "1" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "NoAutoUpdate" /t REG_DWORD /d "0" /f
@REM Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "AUOptions" /t REG_DWORD /d "4" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "AUOptions" /t REG_DWORD /d "2" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "AutomaticMaintenanceEnabled" /t REG_DWORD /d "1" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "ScheduledInstallDay" /t REG_DWORD /d "0" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "ScheduledInstallTime" /t REG_DWORD /d "3" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "AllowMUUpdateService" /t REG_DWORD /d "1" /f

:ENDE

TIMEOUT /T 20

Exit