@ECHO OFF

CD \
ECHO ----------------------------------------------------------------
ECHO MY BAT PROGRAM -- COMMAND LINE USER 
ECHO ----------------------------------------------------------------
ECHO C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\VBS 32-COPIER_SYNC.VBS
ECHO NET_SHARE\VBS 32-COPIER_SYNC.VBS
ECHO HAS SOME CODE TO COPY THESE BATTERY REPORT FILE FOUND AFTER CREATE 
ECHO IN WINDOWS\SYSTEM32
ECHO AND MAKE THEM MOVER AND TIME DATED
ECHO ----------------------------------------------------------------
ECHO.
ECHO ----------------------------------------------------------------
ECHO ----------------------------------------------------------------

@ECHO ON

@ECHO.
powercfg /energy
@ECHO.

@ECHO ----------------------------------------------------------------
@ECHO ----------------------------------------------------------------

@ECHO.
powercfg /batteryreport
@ECHO.
@ECHO ----------------------------------------------------------------
@ECHO ----------------------------------------------------------------

@ECHO OFF

ECHO.

ECHO ----------------------------------------------------------------
ECHO ----------------------------------------------------------------
ECHO ----
ECHO POWERCFG: How To Check Battery Health Report And Energy Report In Windows Using CMD
ECHO https://fossbytes.com/powercfg-check-battery-capacity-battery-health-report-windows-using-cmd/
ECHO ----
ECHO ----------------------------------------------------------------
ECHO ----------------------------------------------------------------

PAUSE
