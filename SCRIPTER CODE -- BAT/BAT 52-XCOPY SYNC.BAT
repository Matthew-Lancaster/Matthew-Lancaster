@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ ALL BATCH FILE COMMAND SETS 
@REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
@REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
@REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
@REM __ ADD CODE IF REQUIRE
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 05-Oct-2018 19:57:00 Creation Time Or Before
@REM __ TO   ---- Sun 07-Oct-2018 08:39:00 Last Update Header Info
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 001 SESSION -- UPDATES HAPPEN ALL THE TIME
@REM ----------------------------------------------------------------
@REM -- GOOD WHEN WANT QUICK XCOPY SYNC BY DATE OLDER NEWER
@REM -- WON'T DO IF PATH IS HIGHER THAN THE COPY FILE OR FILES 
@REM -- CYCLIC COPY ERROR
@REM -- 
@REM -- GOOD AS SCAN EACH DRIVE FOR VOLUME LABEL
@REM -- BUT WAKES UP EACH DRIVE TO DO IT
@REM -- 
@REM -- I GOT A SCANNER IN AUTOHOTKEYS THAT RUN ANYTHING OUT FOR CAMERA 
@REM -- IT SEEM DO BETTER VOLUME LABEL SCAN WITHOUT WAKE UP
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 002 SESSION -- UPDATES HAPPEN ALL THE TIME
@REM ----------------------------------------------------------------
@REM __ FROM ---- Sun 07-Oct-2018 08:31:50
@REM __ TO   ---- Sun 07-Oct-2018 09:08:00 __ Update To Multi On
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- HERE BAT IS NOW SYNC IN 3 LOCATION YET IS SOME ABOUT
@REM -- CENTRALIZING MY BEST CODE
----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 52-XCOPY SYNC.BAT@REM ----------------------------------------------------------------
@REM -- *:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO\*.*
@REM -- D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\*.*
@REM ----------------------------------------------------------------
@REM -- D:\DSC\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\*.*
@REM -- D:\VI_ DSC ME 01\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\*.*
@REM -- *:\INFO _ DOUBLE SCREEN CAMERA\*.* 
@REM ----------------------------------------------------------------
@REM -- IS A SYNCER _ SYNCED BY HERE :) 
@REM -- TO KEEP CENTRAL 
@REM -- MULTI JOB IN ONE
@REM ----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 32-COPIER_SYNC.VBS
@REM ----------------------------------------------------------------

@CD \

@REM __ FORWARD AND THEN REVERSE ------------------------------------
@REM ----------------------------------------------------------------
XCOPY "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT" "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 29-NET_SHARE_ALL.BAT" /D /E /C /Y

XCOPY "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 29-NET_SHARE_ALL.BAT" "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT" /D /E /C /Y
@REM ----------------------------------------------------------------

@ECHO.

@REM ----------------------------------------------------------------
XCOPY "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\*.*" "D:\#0 1 INSTALLATIONS\00 INSTALLER_SET\# 00 Install Progs\# Installed Now\#00 Paid For\GOPRO\INFO _ DOUBLE SCREEN CAMERA FAKE REPLICA GOPRO YELLOW\" /D /E /C /Y

XCOPY "D:\#0 1 INSTALLATIONS\00 INSTALLER_SET\# 00 Install Progs\# Installed Now\#00 Paid For\GOPRO\INFO _ DOUBLE SCREEN CAMERA FAKE REPLICA GOPRO YELLOW\*.*" "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\" /D /E /C /Y
@REM ----------------------------------------------------------------

@ECHO.

PAUSE


@ECHO OFF

GOTO ENDE

setlocal EnableDelayedExpansion


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			IF EXIST "%%G:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO" (
				@ECHO ON
				@ECHO FOUND AT %%G
				@ECHO.
				@REM __ FORWARD AND THEN REVERSE --------------------
				@REM __ ---------------------------------------------
				XCOPY "%%G:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO\*.*" "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\" /D /E /C /Y

				XCOPY "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\*.*" "%%G:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO\" /D /E /C /Y
				@ECHO.
				@ECHO OFF
			)
		)
	)
)
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			IF EXIST "%%G:\INFO _ DOUBLE SCREEN CAMERA" (
				@ECHO ON
				@ECHO FOUND AT %%G
				@ECHO.
				@REM __ ONE WAY AND THEN THE REVERSE ----------------
				@REM __ ---------------------------------------------
				XCOPY "%%G:\INFO _ DOUBLE SCREEN CAMERA\*.*" "D:\DSC\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\" /D /E /C /Y

				XCOPY "D:\DSC\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\*.*" "%%G:\INFO _ DOUBLE SCREEN CAMERA\" /D /E /C /Y
				@ECHO.
				@ECHO OFF
			)
		)
	)
)
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


GOTO VERY_ENDE

::
:IsDeviceReady
setlocal
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF
::
:GetLabel
setlocal
set "label_="
for /f "tokens=5*" %%a in (
  'vol "%~1"^|find "Volume in drive "') do (
    set label_=%%b)
endlocal & set "%2=%label_%" & goto :EOF




@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@REM ----------------------------------------------------------------------------
@REM -- THE END -----------------------------------------------------------------
@REM -- BYE BYE -----------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\



:VERY_ENDER_BLOCK

:VERY_ENDE

:ENDE

EXIT
