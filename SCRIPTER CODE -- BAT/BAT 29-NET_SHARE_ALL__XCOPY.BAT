@CD \

@REM __ FORWARD AND THEN REVERSE ------------------------------------
XCOPY "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT" "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 29-NET_SHARE_ALL.BAT" /D /E /C /Y

XCOPY "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 29-NET_SHARE_ALL.BAT" "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT" /D /E /C /Y

@ECHO.

PAUSE


@ECHO OFF

GOTO ENDE

setlocal EnableDelayedExpansion
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
				
				XCOPY "%%G:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO\*.*" "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\" /D /E /C /Y

				XCOPY "D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\*.*" "%%G:\INFO _ HD DV SPORT YELLOW CAMERA FAKE REPLICA GOPRO\" /D /E /C /Y
				
				@ECHO.
				@ECHO OFF
			)
		)
	)
)


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

