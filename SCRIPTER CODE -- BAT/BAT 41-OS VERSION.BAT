

ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10


ECHO %VARLATCH2%
PAUSE
