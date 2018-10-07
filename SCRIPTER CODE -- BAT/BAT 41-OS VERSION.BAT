@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- --- 28-May-2018 14:05:00 Created Time
@REM __ TO   ---- Sun 07-Oct-2018 04:45:37 Last Editor _ Update Time+Header Info
@REM ----------------------------------------------------------------


ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10


ECHO %VARLATCH2%
PAUSE
