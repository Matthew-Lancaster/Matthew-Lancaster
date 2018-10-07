@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Wed 16-May-2018 02:10:55
@REM __ TO   ---- Wed 16-May-2018 03:32:00 __ 1 HOUR 22 MINUTE _ ONTO NEXT VBSCRIPT 
@REM __ TO   ---- Wed 16-May-2018 04:38:00 __ 2 HOUR 38 MINUTE
@REM __ TO   ---- Wed 16-May-2018 05:05:00 __ 3 HOUR _ LAST BIT 
@REM ----------------------------------------------------------------

@REM ---------------------
@REM HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers
@REM ---------------------

@REM -- THOUGHT I FOUND A REGISTRY KEY SET THE HAD PROGRAM RUNAS ADMIN SETTING
@REM -- BUT A REBOOT THEY GONE


ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10


ECHO %VARLATCH2%
PAUSE
